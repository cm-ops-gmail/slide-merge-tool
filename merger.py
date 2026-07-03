"""
merger.py — Universal PPTX merge logic.

Takes a Root PPTX (content) and a Template PPTX (design) as bytes and returns a
new PPTX where every slide's content from the Root is placed onto the Template's
design (slide master, layouts, theme, background, slide size).

The engine is content-agnostic: nothing about slide order, fonts, titles or
badges is hard-coded. Each shape is deep-copied at the XML level so its original
formatting (fonts, colours, tables, pictures, grouping, effects) is preserved
losslessly, while the Template supplies the surrounding design.

Layout choice can be automatic (best-match) or fully user-driven via a
`layout_map` — a list giving, per Root slide, the index of the Template layout
to use. The UI helpers below expose the template's layouts and a suggested map.
"""

import copy
import io

from pptx import Presentation
from pptx.oxml.ns import qn

# Namespace used by every relationship-id attribute (r:embed, r:link, r:id, ...)
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ─────────────────────────────────────────────
# Layout detection (index-based, so the UI can address layouts by number)
# ─────────────────────────────────────────────

def _blank_layout_index(layouts):
    """Index of the most 'neutral' layout: one named blank, else the one with
    the fewest placeholders (adds the least competing design)."""
    best, best_n = 0, None
    for i, layout in enumerate(layouts):
        if layout.name.strip().lower() == "blank":
            return i
        n = len(layout.placeholders)
        if best_n is None or n < best_n:
            best, best_n = i, n
    return best


def _layout_lookup(prs_out):
    """Return (layouts, name->idx, placeholder_count->idx, blank_idx)."""
    layouts = list(prs_out.slide_layouts)
    by_name, by_phcount = {}, {}
    for i, layout in enumerate(layouts):
        by_name.setdefault(layout.name.strip().lower(), i)
        by_phcount.setdefault(len(layout.placeholders), i)
    return layouts, by_name, by_phcount, _blank_layout_index(layouts)


def _auto_index(root_slide, by_name, by_phcount, blank_idx):
    """Best-match Template layout index for a Root slide.

    1. Same layout name (case-insensitive).
    2. Same placeholder count as the Root slide's own layout.
    3. The neutral / blank layout.
    """
    try:
        name = root_slide.slide_layout.name.strip().lower()
        if name in by_name:
            return by_name[name]
    except Exception:
        pass
    try:
        n = len(root_slide.slide_layout.placeholders)
        if n in by_phcount:
            return by_phcount[n]
    except Exception:
        pass
    return blank_idx


# ─────────────────────────────────────────────
# UI helpers (introspection — no mutation)
# ─────────────────────────────────────────────

def get_template_layout_names(template_bytes: bytes):
    """List the Template's layout names, indexed as the UI/`layout_map` expect."""
    prs = Presentation(io.BytesIO(template_bytes))
    return [layout.name for layout in prs.slide_layouts]


def get_root_slide_previews(root_bytes: bytes):
    """Short text preview + current layout name for each Root slide."""
    prs = Presentation(io.BytesIO(root_bytes))
    previews = []
    for slide in prs.slides:
        text = ""
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                text = " ".join(shape.text_frame.text.split())[:70]
                break
        try:
            layout_name = slide.slide_layout.name
        except Exception:
            layout_name = ""
        previews.append({"text": text or "(no text)", "layout": layout_name})
    return previews


def suggest_layout_map(root_bytes: bytes, template_bytes: bytes):
    """Auto-matched Template layout index for every Root slide (UI defaults)."""
    prs_root = Presentation(io.BytesIO(root_bytes))
    prs_out = Presentation(io.BytesIO(template_bytes))
    _, by_name, by_phcount, blank_idx = _layout_lookup(prs_out)
    return [_auto_index(s, by_name, by_phcount, blank_idx) for s in prs_root.slides]


# ─────────────────────────────────────────────
# Lossless shape copying
# ─────────────────────────────────────────────

def _remap_relationships(src_part, dst_part, new_el):
    """Any cloned shape may reference images/media/hyperlinks by relationship id.
    Those ids are local to the source slide part, so recreate each relationship
    on the destination part and rewrite the id in the copied XML."""
    cache = {}
    for el in new_el.iter():
        for attr in list(el.attrib):
            if not attr.startswith("{%s}" % _R_NS):
                continue
            old_rid = el.get(attr)
            if not old_rid or old_rid not in src_part.rels:
                continue
            if old_rid not in cache:
                rel = src_part.rels[old_rid]
                if rel.is_external:
                    cache[old_rid] = dst_part.relate_to(
                        rel.target_ref, rel.reltype, is_external=True
                    )
                else:
                    cache[old_rid] = dst_part.relate_to(rel.target_part, rel.reltype)
            el.set(attr, cache[old_rid])


def _insert_shape(dst_slide, new_el):
    """Append a cloned shape element into the slide's shape tree, keeping it
    before any trailing <p:extLst> so the file stays schema-valid."""
    spTree = dst_slide.shapes._spTree
    extLst = spTree.find(qn("p:extLst"))
    if extLst is not None:
        extLst.addprevious(new_el)
    else:
        spTree.append(new_el)


def _scale_shape(src_shape, dst_shape, sx, sy):
    """When Root and Template slide sizes differ, rescale a top-level shape so it
    still fits the Template's coordinate space."""
    for attr, factor in (("left", sx), ("top", sy), ("width", sx), ("height", sy)):
        try:
            v = getattr(src_shape, attr)
            if v is not None:
                setattr(dst_shape, attr, int(round(v * factor)))
        except Exception:
            pass


def _reassign_shape_ids(dst_slide):
    """Give every shape a unique cNvPr id within the slide (cloning can duplicate
    ids, which makes PowerPoint prompt to 'repair')."""
    idx = 1
    for cNvPr in dst_slide.shapes._spTree.iter(qn("p:cNvPr")):
        cNvPr.set("id", str(idx))
        idx += 1


# ─────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────

def merge_presentations(root_bytes: bytes, template_bytes: bytes, layout_map=None):
    """
    Apply the Template's design to the Root's content.

    Parameters
    ----------
    root_bytes     : bytes of the Root PPTX (content source)
    template_bytes : bytes of the Template PPTX (design source)
    layout_map     : optional list — layout_map[i] is the Template layout index
                     for Root slide i. Entries that are None (or out of range),
                     and slides beyond the list, fall back to auto-matching.
                     Pass None to auto-match every slide.

    Returns
    -------
    (bytes of the new PPTX, list[str] of non-fatal warnings)
    """
    prs_root = Presentation(io.BytesIO(root_bytes))
    prs_out = Presentation(io.BytesIO(template_bytes))

    # Keep the Template's slide size; scale Root content if the sizes differ.
    tw, th = prs_out.slide_width, prs_out.slide_height
    rw, rh = prs_root.slide_width, prs_root.slide_height
    sx = (tw / rw) if rw else 1.0
    sy = (th / rh) if rh else 1.0
    need_scale = abs(sx - 1) > 0.01 or abs(sy - 1) > 0.01

    # ── Wipe the Template's own slides (we only want its design) ────────────────
    sldIdLst = prs_out.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs_out.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)

    layouts, by_name, by_phcount, blank_idx = _layout_lookup(prs_out)

    warnings = []

    for slide_idx, root_slide in enumerate(prs_root.slides):
        # ── Pick the Template layout ──────────────────────────────────────────
        chosen = None
        if layout_map is not None and slide_idx < len(layout_map):
            chosen = layout_map[slide_idx]
        if isinstance(chosen, int) and 0 <= chosen < len(layouts):
            layout = layouts[chosen]
        else:
            layout = layouts[_auto_index(root_slide, by_name, by_phcount, blank_idx)]

        new_slide = prs_out.slides.add_slide(layout)

        # ── Copy every shape losslessly ───────────────────────────────────────
        for shape in root_slide.shapes:
            try:
                new_el = copy.deepcopy(shape._element)
                _remap_relationships(root_slide.part, new_slide.part, new_el)
                _insert_shape(new_slide, new_el)
                if need_scale:
                    _scale_shape(shape, new_slide.shapes[-1], sx, sy)
            except Exception as e:
                warnings.append(
                    f"Slide {slide_idx + 1} | shape '{getattr(shape, 'name', '?')}': {e}"
                )

        _reassign_shape_ids(new_slide)

    out = io.BytesIO()
    prs_out.save(out)
    return out.getvalue(), warnings
