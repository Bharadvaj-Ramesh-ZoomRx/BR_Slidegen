"""Smoke test for stamp_refresh_dynamic_tags (per spec §16.3).

The tool that drives refresh must rewrite LASTREFRESHTIME on success and
REFRESHERRORMSG on failure on every refreshed shape.  This test verifies
that the function locates the shape's tag file via slide rels, replaces
any pre-existing dynamic tags, and writes back a valid tagLst.
"""
import sys
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))


def _build_minimal_pptx_with_tag(tmp_path: Path) -> Path:
    """Build a tiny .pptx with one slide carrying one shape that has a
    custDataLst → tag file referenced via rels.  Just enough for
    stamp_refresh_dynamic_tags to find and modify."""
    pptx = tmp_path / "smoke.pptx"

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/tags/tag1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"/>
</Types>"""

    pkg_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

    pres_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"""

    slide_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree>'
        '<p:sp>'
        '<p:nvSpPr>'
        '<p:cNvPr id="2" name="TestShape"/>'
        '<p:cNvSpPr/>'
        '<p:nvPr><p:custDataLst><p:tags r:id="rId2"/></p:custDataLst></p:nvPr>'
        '</p:nvSpPr>'
        '<p:spPr/></p:sp>'
        '</p:spTree></p:cSld>'
        '</p:sld>'
    )

    slide_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" Target="../tags/tag1.xml"/>
</Relationships>"""

    tag_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:tag name="REPORTCONFIGHASH" val="abc123"/>'
        '</p:tagLst>'
    )

    with zipfile.ZipFile(pptx, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("ppt/presentation.xml", pres_xml)
        z.writestr("ppt/slides/slide1.xml", slide_xml)
        z.writestr("ppt/slides/_rels/slide1.xml.rels", slide_rels)
        z.writestr("ppt/tags/tag1.xml", tag_xml)

    return pptx


def _read_tag_xml(pptx: Path, name: str) -> str:
    with zipfile.ZipFile(pptx) as z:
        return z.read(name).decode("utf-8")


def test_stamp_dynamic_tags_writes_lastrefreshtime_on_success(tmp_path):
    from slidegen.intelligent_refresh import stamp_refresh_dynamic_tags

    pptx = _build_minimal_pptx_with_tag(tmp_path)
    refresh_results = {
        "slides": [{
            "slide_index": 0,
            "charts": [{"name": "TestShape", "status": "ok"}],
            "tables": [],
        }],
    }

    stamps = stamp_refresh_dynamic_tags(str(pptx), refresh_results)
    assert stamps == {(0, "TestShape"): "stamped"}

    after = _read_tag_xml(pptx, "ppt/tags/tag1.xml")
    assert "LASTREFRESHTIME" in after
    # On success, REFRESHERRORMSG must NOT be present
    assert "REFRESHERRORMSG" not in after
    # Original tag should still be there
    assert "REPORTCONFIGHASH" in after


def test_stamp_dynamic_tags_writes_error_on_failure(tmp_path):
    from slidegen.intelligent_refresh import stamp_refresh_dynamic_tags

    pptx = _build_minimal_pptx_with_tag(tmp_path)
    refresh_results = {
        "slides": [{
            "slide_index": 0,
            "charts": [{
                "name": "TestShape",
                "status": "alignment_failed",
                "error": "alignment_failed: no API series matched source",
            }],
            "tables": [],
        }],
    }

    stamp_refresh_dynamic_tags(str(pptx), refresh_results)
    after = _read_tag_xml(pptx, "ppt/tags/tag1.xml")
    assert "LASTREFRESHTIME" in after
    assert "REFRESHERRORMSG" in after
    assert "alignment_failed" in after


def test_stamp_dynamic_tags_clears_stale_error_on_success(tmp_path):
    """A shape that previously errored must have its REFRESHERRORMSG
    deleted on the next successful refresh (per §16.3)."""
    from slidegen.intelligent_refresh import stamp_refresh_dynamic_tags

    pptx = _build_minimal_pptx_with_tag(tmp_path)

    # First, simulate a failed refresh
    stamp_refresh_dynamic_tags(str(pptx), {
        "slides": [{
            "slide_index": 0,
            "charts": [{"name": "TestShape", "status": "error",
                        "error": "stale failure"}],
            "tables": [],
        }],
    })
    after_fail = _read_tag_xml(pptx, "ppt/tags/tag1.xml")
    assert "REFRESHERRORMSG" in after_fail

    # Now run a successful refresh — REFRESHERRORMSG must be cleared
    stamp_refresh_dynamic_tags(str(pptx), {
        "slides": [{
            "slide_index": 0,
            "charts": [{"name": "TestShape", "status": "ok"}],
            "tables": [],
        }],
    })
    after_ok = _read_tag_xml(pptx, "ppt/tags/tag1.xml")
    assert "REFRESHERRORMSG" not in after_ok
    assert "stale failure" not in after_ok


def test_stamp_dynamic_tags_no_op_on_empty_results(tmp_path):
    from slidegen.intelligent_refresh import stamp_refresh_dynamic_tags

    pptx = _build_minimal_pptx_with_tag(tmp_path)
    stamps = stamp_refresh_dynamic_tags(str(pptx), {"slides": []})
    assert stamps == {}


def test_stamp_dynamic_tags_idempotent_across_runs(tmp_path):
    """Running the stamper twice with identical results must not duplicate
    LASTREFRESHTIME entries — only the latest stays."""
    from slidegen.intelligent_refresh import stamp_refresh_dynamic_tags

    pptx = _build_minimal_pptx_with_tag(tmp_path)
    refresh_results = {
        "slides": [{
            "slide_index": 0,
            "charts": [{"name": "TestShape", "status": "ok"}],
            "tables": [],
        }],
    }

    stamp_refresh_dynamic_tags(str(pptx), refresh_results)
    stamp_refresh_dynamic_tags(str(pptx), refresh_results)

    after = _read_tag_xml(pptx, "ppt/tags/tag1.xml")
    # Exactly one LASTREFRESHTIME, not two
    assert after.count("LASTREFRESHTIME") == 1
