"""
Unit tests for data extraction functions and config validation.

Covers:
- extract_by_question_code: zero-value handling, basic extraction
- pct/straight converters: edge cases
- make_label_shortener: brand replacements, truncation
- parse_color: valid and invalid hex strings
- ProjectConfig.validate: sample sizes, sort_by, brand refs
"""

import os
import sys
import pytest
import pandas as pd

BASE = os.path.dirname(__file__)
ROOT = os.path.dirname(BASE)
sys.path.insert(0, ROOT)

from slidegen.pipeline.data_loaders import (
    pct, straight, delta,
    extract_by_question_code,
    extract_multi_question_code,
    make_label_shortener,
)
from slidegen.pipeline.project_config import parse_color, load_project_config


# ── Value converters ────────────────────────────────────────────────────────

class TestConverters:
    def test_pct_normal(self):
        assert pct(0.5) == 50.0
        assert pct(0.123) == 12.3
        assert pct(1.0) == 100.0

    def test_pct_zero(self):
        """Zero is a valid value, not None."""
        assert pct(0) == 0.0
        assert pct(0.0) == 0.0

    def test_pct_none(self):
        assert pct(None) is None

    def test_pct_non_numeric(self):
        assert pct("abc") is None
        assert pct("") is None

    def test_straight_normal(self):
        assert straight(50.0) == 50.0
        assert straight(12.34) == 12.3

    def test_straight_zero(self):
        assert straight(0) == 0.0

    def test_straight_none(self):
        assert straight(None) is None

    def test_delta_normal(self):
        assert delta(50.0, 40.0) == 10.0
        assert delta(30.0, 40.0) == -10.0

    def test_delta_zero(self):
        assert delta(0.0, 0.0) == 0.0

    def test_delta_none(self):
        assert delta(None, 50.0) is None
        assert delta(50.0, None) is None


# ── extract_by_question_code ────────────────────────────────────────────────

def _make_test_sheet(codes_and_values):
    """Build a minimal DataFrame mimicking a survey sheet.

    codes_and_values: list of (code, desc, prior_val, current_val) tuples.
    First element should be the question code header row.
    """
    rows = []
    for code, desc, prior, current in codes_and_values:
        rows.append([code, desc, prior, current])
    return pd.DataFrame(rows)


class TestExtractByQuestionCode:
    def test_basic_extraction(self):
        sheet = _make_test_sheet([
            ("Q1_10Z", "Message Recall", None, None),  # header row
            ("", "Msg A", 0.5, 0.7),                   # sub-row 1
            ("", "Msg B", 0.3, 0.4),                   # sub-row 2
        ])
        result = extract_by_question_code(sheet, "Q1_10Z", q_prior_col=2, q_current_col=3)
        assert len(result) == 2
        assert result[0]["desc"] == "Msg A"
        assert result[0]["prior"] == 50.0
        assert result[0]["current"] == 70.0
        assert result[1]["desc"] == "Msg B"

    def test_zero_values_preserved(self):
        """Regression test: rows with 0% values should NOT be dropped."""
        sheet = _make_test_sheet([
            ("Q1_10Z", "Recall", None, None),
            ("", "Msg A", 0.0, 0.0),     # both zero — should be kept
            ("", "Msg B", 0.0, 0.5),     # prior zero — should be kept
            ("", "Msg C", 0.5, 0.0),     # current zero — should be kept
        ])
        result = extract_by_question_code(sheet, "Q1_10Z", q_prior_col=2, q_current_col=3)
        assert len(result) == 3, f"Expected 3 rows, got {len(result)}: zero values should be preserved"
        assert result[0]["prior"] == 0.0
        assert result[0]["current"] == 0.0

    def test_both_none_skipped(self):
        """Rows where both prior and current are None should be skipped."""
        sheet = _make_test_sheet([
            ("Q1_10Z", "Recall", None, None),
            ("", "Msg A", None, None),    # both None — should be skipped
            ("", "Msg B", 0.5, 0.7),      # valid — should be kept
        ])
        result = extract_by_question_code(sheet, "Q1_10Z", q_prior_col=2, q_current_col=3)
        # Msg A has current_val=None, so pd.notna(q_current) is False → skipped by line 109 check
        # Only Msg B should appear
        assert len(result) == 1
        assert result[0]["desc"] == "Msg B"

    def test_code_not_found(self):
        sheet = _make_test_sheet([
            ("Q1_10Z", "Recall", None, None),
            ("", "Msg A", 0.5, 0.7),
        ])
        result = extract_by_question_code(sheet, "NONEXISTENT", q_prior_col=2, q_current_col=3)
        assert result == []

    def test_base_rows_excluded(self):
        """Rows starting with 'Base' should be excluded."""
        sheet = _make_test_sheet([
            ("Q1_10Z", "Recall", None, None),
            ("", "Base (n=100)", 100, 100),   # base row — excluded
            ("", "Msg A", 0.5, 0.7),
        ])
        result = extract_by_question_code(sheet, "Q1_10Z", q_prior_col=2, q_current_col=3)
        assert len(result) == 1
        assert result[0]["desc"] == "Msg A"

    def test_label_fn_applied(self):
        sheet = _make_test_sheet([
            ("Q1_10Z", "Recall", None, None),
            ("", "Very Long Message Description", 0.5, 0.7),
        ])
        result = extract_by_question_code(
            sheet, "Q1_10Z", q_prior_col=2, q_current_col=3,
            label_fn=lambda s: s[:10]
        )
        assert result[0]["desc"] == "Very Long "

    def test_straight_converter(self):
        """When pct_mode=straight, values pass through without *100."""
        sheet = _make_test_sheet([
            ("Q1_10Z", "Recall", None, None),
            ("", "Msg A", 50.0, 70.0),
        ])
        result = extract_by_question_code(
            sheet, "Q1_10Z", q_prior_col=2, q_current_col=3,
            converter=straight
        )
        assert result[0]["prior"] == 50.0
        assert result[0]["current"] == 70.0


# ── make_label_shortener ────────────────────────────────────────────────────

class TestLabelShortener:
    def test_keyword_match(self):
        shortcuts = [{"keywords": ["recall", "message"], "short": "MR"}]
        fn = make_label_shortener(shortcuts)
        assert fn("Total Message Recall") == "MR"

    def test_no_match_passthrough(self):
        fn = make_label_shortener([])
        assert fn("Short label") == "Short label"

    def test_truncation(self):
        fn = make_label_shortener([], max_len=10)
        assert fn("A very long label that exceeds") == "A very lo\u2026"

    def test_brand_replacements(self):
        replacements = {
            "Acme Corp (Formerly BigCo)": "Acme",
            "[COMPANY]": "Acme",
            "[PRODUCT]": "Widget",
        }
        fn = make_label_shortener([], brand_replacements=replacements)
        assert fn("[COMPANY] product line") == "Acme product line"
        assert fn("[PRODUCT] awareness") == "Widget awareness"
        assert fn("Acme Corp (Formerly BigCo) total") == "Acme total"

    def test_brand_replacements_none(self):
        """None brand_replacements should not crash."""
        fn = make_label_shortener([], brand_replacements=None)
        assert fn("Test label") == "Test label"


# ── parse_color ─────────────────────────────────────────────────────────────

class TestParseColor:
    def test_valid_hex_with_hash(self):
        c = parse_color("#FF5824")
        assert c == (0xFF, 0x58, 0x24)

    def test_valid_hex_without_hash(self):
        c = parse_color("7030A0")
        assert c == (0x70, 0x30, 0xA0)

    def test_invalid_short_hex(self):
        with pytest.raises(ValueError, match="Invalid hex color"):
            parse_color("#FF")

    def test_invalid_chars(self):
        with pytest.raises(ValueError, match="Invalid hex color"):
            parse_color("#ZZZZZZ")

    def test_empty_string(self):
        with pytest.raises(ValueError, match="Invalid hex color"):
            parse_color("")


# ── Config validation ───────────────────────────────────────────────────────

class TestConfigValidation:
    """Test config validation using the test project's config.yaml."""

    @pytest.fixture
    def config(self):
        yaml_path = os.path.join(BASE, "test_project", "config.yaml")
        if not os.path.exists(yaml_path):
            pytest.skip("test_project/config.yaml not found")
        return load_project_config(yaml_path)

    def test_valid_config_has_no_new_errors(self, config):
        """Validate that new validation rules don't produce false positives on test config.

        Note: test config may have intentional gaps (e.g. missing extractions for demo asks).
        We only check that no sample_size or brand errors appear.
        """
        errors = config.validate()
        sample_errors = [e for e in errors if "sample_sizes" in e]
        brand_errors = [e for e in errors if "brand" in e and "not in config.brands" in e]
        assert sample_errors == [], f"Sample size validation errors: {sample_errors}"
        assert brand_errors == [], f"Brand validation errors: {brand_errors}"

    def test_bad_slide_type_detected(self, config):
        # Temporarily modify an ask to have invalid slide_type
        original = config.asks[0].slide_type
        config.asks[0].slide_type = "nonexistent_renderer"
        errors = config.validate()
        assert any("nonexistent_renderer" in e for e in errors)
        config.asks[0].slide_type = original

    def test_bad_data_key_detected(self, config):
        original = config.asks[0].data_key
        config.asks[0].data_key = "nonexistent_extraction"
        errors = config.validate()
        assert any("nonexistent_extraction" in e for e in errors)
        config.asks[0].data_key = original

    def test_negative_sample_size_detected(self, config):
        for key in config.sample_sizes:
            original = config.sample_sizes[key].current
            config.sample_sizes[key].current = -1
            errors = config.validate()
            assert any("current must be positive" in e for e in errors)
            config.sample_sizes[key].current = original
            break

    def test_bad_brand_ref_detected(self, config):
        if config.asks:
            original = config.asks[0].brand
            config.asks[0].brand = "nonexistent_brand"
            errors = config.validate()
            assert any("nonexistent_brand" in e for e in errors)
            config.asks[0].brand = original


    def test_duplicate_extraction_id_detected(self, config):
        """Duplicate extraction IDs should produce a validation error."""
        from slidegen.pipeline.project_config import DataExtractionConfig
        original_extractions = list(config.extractions)
        # Add a duplicate of the first extraction
        dup = DataExtractionConfig(
            id=config.extractions[0].id,
            method="mock",
            sheet="",
            params={"rows": []},
        )
        config.extractions.append(dup)
        errors = config.validate()
        assert any("duplicate extraction ID" in e for e in errors), \
            f"Expected duplicate ID error, got: {errors}"
        config.extractions[:] = original_extractions

    def test_unknown_method_detected(self, config):
        """Unknown extraction method should produce a validation error."""
        from slidegen.pipeline.project_config import DataExtractionConfig
        original_extractions = list(config.extractions)
        config.extractions.append(DataExtractionConfig(
            id="bad_method_test",
            method="nonexistent_method",
            sheet="primary",
            params={},
        ))
        errors = config.validate()
        assert any("nonexistent_method" in e for e in errors)
        config.extractions[:] = original_extractions

    def test_missing_required_params_detected(self, config):
        """Extraction without required params should produce a validation error."""
        from slidegen.pipeline.project_config import DataExtractionConfig
        original_extractions = list(config.extractions)
        # question_code requires params.code
        config.extractions.append(DataExtractionConfig(
            id="no_code_test",
            method="question_code",
            sheet="primary",
            params={},
        ))
        errors = config.validate()
        assert any("requires params.code" in e for e in errors)
        config.extractions[:] = original_extractions

    def test_valid_config_passes(self, config):
        """The test project config should pass validation cleanly."""
        errors = config.validate()
        assert errors == [], f"Expected no errors, got: {errors}"


class TestValidateCLI:
    """Test the validate CLI command."""

    def test_validate_valid_config(self):
        import subprocess
        yaml_path = os.path.join(BASE, "test_project", "config.yaml")
        result = subprocess.run(
            [sys.executable, "-m", "slidegen", "validate", yaml_path],
            capture_output=True, text=True, cwd=ROOT,
        )
        assert result.returncode == 0
        assert "Config is valid" in result.stdout

    def test_validate_missing_file(self):
        import subprocess
        result = subprocess.run(
            [sys.executable, "-m", "slidegen", "validate", "nonexistent.yaml"],
            capture_output=True, text=True, cwd=ROOT,
        )
        assert result.returncode == 1
        assert "not found" in result.stdout

    def test_validate_no_args(self):
        import subprocess
        result = subprocess.run(
            [sys.executable, "-m", "slidegen", "validate"],
            capture_output=True, text=True, cwd=ROOT,
        )
        assert result.returncode == 1

    def test_unknown_command(self):
        import subprocess
        result = subprocess.run(
            [sys.executable, "-m", "slidegen", "bogus_command"],
            capture_output=True, text=True, cwd=ROOT,
        )
        assert result.returncode == 1
        assert "Unknown command" in result.stdout


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
