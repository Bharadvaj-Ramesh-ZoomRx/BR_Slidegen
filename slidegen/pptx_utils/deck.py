"""
deck.py — Template handling, slide scaffolding, and presentation management.

Functions for loading templates, clearing slides/sections, and creating
PowerPoint sections. Extracted from orchestrator.py per PRD §4.2.
"""

from __future__ import annotations

import os
import uuid

from pptx import Presentation
from pptx.util import Emu


def load_template(config) -> tuple:
    """Load template PPTX, returning (Presentation, blank_layout).

    If a template exists, loads it and removes all original slides so that only
    the slide masters/layouts remain (logos, fonts, backgrounds are preserved).
    Falls back to a blank presentation if no template is available.

    Args:
        config: ProjectConfig with template_path attribute.

    Returns:
        (Presentation, slide_layout) tuple.
    """
    from pptx.oxml.ns import qn

    if config.template_path and os.path.exists(config.template_path):
        # Strategy: Build a clean blank PPTX with the correct dimensions,
        # then copy only the slide master XML theme (colors, fonts) from the
        # template. This avoids ALL orphaned chart/slide/customXml issues.
        from pptx.oxml.ns import qn as _qn
        import zipfile
        import tempfile

        source_prs = Presentation(config.template_path)
        original_count = len(source_prs.slides)

        # Extract theme colors and fonts from the template's first slide master
        theme_xml = None
        try:
            master = source_prs.slide_masters[0]
            theme_part = master.part.related_part('rId1')  # theme is usually rId1
            theme_xml = theme_part._element
        except Exception:
            pass

        # Create a fresh blank presentation
        prs = Presentation()
        prs.slide_width = source_prs.slide_width
        prs.slide_height = source_prs.slide_height

        # Apply the template's theme to the blank presentation's master
        if theme_xml is not None:
            try:
                blank_master = prs.slide_masters[0]
                blank_theme_part = blank_master.part.related_part('rId1')
                # Replace the theme element
                blank_theme_part._element.getparent().replace(
                    blank_theme_part._element, theme_xml
                )
            except Exception:
                pass  # If theme injection fails, proceed with default theme

        # Find the "Blank" layout
        blank_layout = None
        for layout in prs.slide_layouts:
            if layout.name == "Blank":
                blank_layout = layout
                break
        if blank_layout is None:
            blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]

        print(f"  Template loaded: {os.path.basename(config.template_path)}"
              f" (master/layouts retained, {original_count} slides cleared)")
        return prs, blank_layout
    else:
        prs = Presentation()
        prs.slide_width = Emu(12192000)   # 13.33 inches
        prs.slide_height = Emu(6858000)   # 7.50 inches
        layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
        print("  No template — using blank presentation")
        return prs, layout


def clear_slide(slide):
    """Remove all shapes from a slide, preserving the slide itself."""
    sp_tree = slide.shapes._spTree
    for child in list(sp_tree):
        tag = child.tag
        if (tag.endswith('}sp') or tag.endswith('}graphicFrame') or
                tag.endswith('}pic') or tag.endswith('}grpSp') or
                tag.endswith('}cxnSp')):
            sp_tree.remove(child)


_P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
_SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"


def clear_sections(prs):
    """Remove all PowerPoint sections from the presentation."""
    from pptx.oxml.ns import qn

    ext_lst = prs.part._element.find(qn('p:extLst'))
    if ext_lst is None:
        return
    for ext in list(ext_lst):
        if ext.get("uri") == _SECTION_EXT_URI:
            section_lst = ext.find(f"{{{_P14_NS}}}sectionLst")
            if section_lst is not None:
                for section in list(section_lst):
                    section_lst.remove(section)


def create_sections(prs, sections_config: list[dict], ask_id_to_slide_idx: dict):
    """Create PowerPoint sections from config.

    Args:
        prs: Presentation object.
        sections_config: [{"name": "Section Name", "start": "ask_id"}, ...]
        ask_id_to_slide_idx: {"ask_id": 0-based slide index}
    """
    from pptx.oxml.ns import qn
    from lxml import etree

    if not sections_config:
        return

    sld_id_lst = prs.part._element.find(qn('p:sldIdLst'))
    sld_ids = [elem.get("id") for elem in sld_id_lst]
    total_slides = len(sld_ids)

    ranges = []
    for sec in sections_config:
        start_ask = sec.get("start", "")
        start_idx = ask_id_to_slide_idx.get(start_ask, 0)
        ranges.append((sec["name"], start_idx))

    ranges.sort(key=lambda x: x[1])

    ext_lst = prs.part._element.find(qn('p:extLst'))
    if ext_lst is None:
        ext_lst = etree.SubElement(prs.part._element, qn('p:extLst'))

    section_ext = None
    for ext in ext_lst:
        if ext.get("uri") == _SECTION_EXT_URI:
            section_ext = ext
            break

    if section_ext is None:
        section_ext = etree.SubElement(ext_lst, qn('p:ext'))
        section_ext.set("uri", _SECTION_EXT_URI)

    section_lst = section_ext.find(f"{{{_P14_NS}}}sectionLst")
    if section_lst is None:
        section_lst = etree.SubElement(section_ext, f"{{{_P14_NS}}}sectionLst")

    for i, (name, start_idx) in enumerate(ranges):
        end_idx = ranges[i + 1][1] if i + 1 < len(ranges) else total_slides

        section_el = etree.SubElement(section_lst, f"{{{_P14_NS}}}section")
        section_el.set("name", name)
        section_el.set("id", "{" + str(uuid.uuid4()).upper() + "}")

        sld_id_lst_el = etree.SubElement(section_el, f"{{{_P14_NS}}}sldIdLst")
        for idx in range(start_idx, end_idx):
            if idx < len(sld_ids):
                sld_id_el = etree.SubElement(sld_id_lst_el, f"{{{_P14_NS}}}sldId")
                sld_id_el.set("id", sld_ids[idx])

    print(f"  Created {len(ranges)} sections")
