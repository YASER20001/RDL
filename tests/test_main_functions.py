"""
Comprehensive pytest test suite for KBR RDL Data Harmonizer — Core Engine
Tests all core logic functions used by the Streamlit app.
"""

import sys
import os
import io
import re
import time
import uuid
import pandas as pd
import numpy as np

# Add backend directory to path so we can import core module
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

from core import (
    _safe_str, _normalize_abbreviations, _tokenize, _has_word_overlap,
    _expand_compound_names, fuzzy_match, classify_match, ABBREVIATIONS,
    read_aramco, read_cfihos, read_kbr, read_ltc, read_sa_doc,
    read_aramco_attributes, read_ltc_attributes,
    run_harmonization, build_export_df, get_demo_data,
)


# ══════════════════════════════════════════════════════════════════════
# Helper: Create test Excel files
# ══════════════════════════════════════════════════════════════════════
def make_excel(data, sheet_name="Sheet1"):
    """Create an in-memory Excel file from a dict."""
    buf = io.BytesIO()
    df = pd.DataFrame(data)
    df.to_excel(buf, sheet_name=sheet_name, index=False)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════
# Tests: _safe_str
# ══════════════════════════════════════════════════════════════════════
class TestSafeStr:
    def test_normal_string(self):
        assert _safe_str("Hello") == "Hello"

    def test_nan_value(self):
        assert _safe_str(float("nan")) == ""

    def test_none_value(self):
        assert _safe_str(None) == ""

    def test_numpy_nan(self):
        assert _safe_str(np.nan) == ""

    def test_string_nan(self):
        assert _safe_str("nan") == ""

    def test_string_NaN(self):
        assert _safe_str("NaN") == ""

    def test_whitespace_stripped(self):
        assert _safe_str("  hello  ") == "hello"

    def test_empty_string(self):
        assert _safe_str("") == ""

    def test_numeric(self):
        assert _safe_str(42) == "42"

    def test_float_value(self):
        assert _safe_str(3.14) == "3.14"


# ══════════════════════════════════════════════════════════════════════
# Tests: _normalize_abbreviations
# ══════════════════════════════════════════════════════════════════════
class TestNormalizeAbbreviations:
    def test_hx_expansion(self):
        result = _normalize_abbreviations("Shell Tube HX")
        assert "heat exchanger" in result

    def test_vlv_expansion(self):
        result = _normalize_abbreviations("Control VLV")
        assert "valve" in result

    def test_pmp_expansion(self):
        result = _normalize_abbreviations("Centrifugal PMP")
        assert "pump" in result

    def test_no_abbreviation(self):
        result = _normalize_abbreviations("Centrifugal Pump")
        assert result == "centrifugal pump"

    def test_multiple_abbreviations(self):
        result = _normalize_abbreviations("ELEC MTR DRV")
        assert "electric" in result
        assert "motor" in result
        assert "driver" in result

    def test_case_insensitive(self):
        result = _normalize_abbreviations("HX")
        assert "heat exchanger" in result

    def test_all_abbreviations_defined(self):
        """Every abbreviation in ABBREVIATIONS dict should expand correctly."""
        for abbr, expansion in ABBREVIATIONS.items():
            result = _normalize_abbreviations(abbr)
            assert expansion in result, f"'{abbr}' should expand to contain '{expansion}', got '{result}'"

    def test_empty_string(self):
        assert _normalize_abbreviations("") == ""


# ══════════════════════════════════════════════════════════════════════
# Tests: _tokenize
# ══════════════════════════════════════════════════════════════════════
class TestTokenize:
    def test_basic_tokenization(self):
        tokens = _tokenize("Centrifugal Pump")
        assert "centrifugal" in tokens
        assert "pump" in tokens

    def test_noise_removal(self):
        tokens = _tokenize("The General System Type")
        # "the", "general", "system", "type" are all noise words
        assert "the" not in tokens
        assert "general" not in tokens
        assert "system" not in tokens
        assert "type" not in tokens

    def test_short_words_excluded(self):
        """Words less than 2 chars are excluded by regex."""
        tokens = _tokenize("A B C Pump")
        assert "a" not in tokens
        assert "pump" in tokens

    def test_numbers_excluded(self):
        """Regex [a-z]{2,} excludes pure numbers."""
        tokens = _tokenize("Pump 123 Model")
        assert "pump" in tokens
        assert "model" in tokens
        assert "123" not in tokens

    def test_empty_string(self):
        assert _tokenize("") == set()


# ══════════════════════════════════════════════════════════════════════
# Tests: _has_word_overlap
# ══════════════════════════════════════════════════════════════════════
class TestHasWordOverlap:
    def test_direct_overlap(self):
        assert _has_word_overlap("Centrifugal Pump", "Centrifugal Pump Unit") is True

    def test_no_overlap(self):
        assert _has_word_overlap("Motor Driver", "Centrifugal Pump") is False

    def test_abbreviation_overlap(self):
        """HX should match Heat Exchanger via abbreviation expansion."""
        assert _has_word_overlap("Shell Tube HX", "Shell and Tube Heat Exchanger") is True

    def test_substring_overlap(self):
        """'compressor' contains 'compress' — should match via substring."""
        assert _has_word_overlap("Compressor", "Reciprocating Compressor") is True

    def test_empty_name_returns_true(self):
        """Empty tokens should return True (permissive)."""
        assert _has_word_overlap("", "Pump") is True

    def test_both_empty(self):
        assert _has_word_overlap("", "") is True

    def test_partial_word_match(self):
        """'pump' in 'pumping' — substring containment >= 3 chars."""
        assert _has_word_overlap("Pump Station", "Pumping Unit") is True


# ══════════════════════════════════════════════════════════════════════
# Tests: _expand_compound_names
# ══════════════════════════════════════════════════════════════════════
class TestExpandCompoundNames:
    def test_no_slash(self):
        candidates = [{"name": "Pump", "id": "1"}]
        names, idx_map = _expand_compound_names(candidates)
        assert names == ["Pump"]
        assert idx_map == [0]

    def test_with_slash(self):
        candidates = [{"name": "Junction Box / Splice Case", "id": "1"}]
        names, idx_map = _expand_compound_names(candidates)
        assert "Junction Box / Splice Case" in names
        assert "Junction Box" in names
        assert "Splice Case" in names
        # All should map back to index 0
        assert all(i == 0 for i in idx_map)

    def test_multiple_candidates(self):
        candidates = [
            {"name": "Pump", "id": "1"},
            {"name": "Valve / Gate Valve", "id": "2"},
        ]
        names, idx_map = _expand_compound_names(candidates)
        assert len(names) == 4  # Pump, Valve / Gate Valve, Valve, Gate Valve
        assert idx_map[0] == 0  # Pump -> 0
        assert idx_map[1] == 1  # Valve / Gate Valve -> 1
        assert idx_map[2] == 1  # Valve -> 1
        assert idx_map[3] == 1  # Gate Valve -> 1


# ══════════════════════════════════════════════════════════════════════
# Tests: fuzzy_match
# ══════════════════════════════════════════════════════════════════════
class TestFuzzyMatch:
    def setup_method(self):
        self.candidates = [
            {"id": "1", "name": "Centrifugal Pump"},
            {"id": "2", "name": "Shell and Tube Heat Exchanger"},
            {"id": "3", "name": "Pressure Vessel"},
            {"id": "4", "name": "Control Valve"},
            {"id": "5", "name": "Electric Motor"},
        ]

    def test_exact_match(self):
        result = fuzzy_match("Centrifugal Pump", self.candidates, 75)
        assert result is not None
        assert result["name"] == "Centrifugal Pump"
        assert result["score"] >= 95

    def test_close_match(self):
        result = fuzzy_match("Centrifugal Pump Unit", self.candidates, 75)
        assert result is not None
        assert result["name"] == "Centrifugal Pump"

    def test_no_match_below_threshold(self):
        result = fuzzy_match("Completely Unrelated Device", self.candidates, 75)
        assert result is None

    def test_abbreviation_match_hx(self):
        """'Shell Tube HX' should match 'Shell and Tube Heat Exchanger'."""
        result = fuzzy_match("Shell Tube HX", self.candidates, 60)
        assert result is not None
        assert "Heat Exchanger" in result["name"]

    def test_token_set_ratio_partial(self):
        """'Storage Tank' should match 'Atmospheric Storage Tank' via token_set_ratio."""
        candidates = [{"id": "1", "name": "Atmospheric Storage Tank"}]
        result = fuzzy_match("Storage Tank", candidates, 75)
        assert result is not None
        assert result["name"] == "Atmospheric Storage Tank"

    def test_compound_name_search(self):
        """Searching for part of a compound name."""
        candidates = [{"id": "1", "name": "Junction Box / Splice Case"}]
        result = fuzzy_match("Junction Box", candidates, 75)
        assert result is not None

    def test_empty_candidates(self):
        result = fuzzy_match("Pump", [], 75)
        assert result is None

    def test_word_overlap_rejection(self):
        """Should reject matches that have no word overlap (false positives)."""
        candidates = [{"id": "1", "name": "Pump"}]
        # "ZZZZ" has no word overlap with "Pump", even if fuzzy score is somehow high
        result = fuzzy_match("ZZZZZZZ", candidates, 0)
        assert result is None

    def test_score_is_integer(self):
        result = fuzzy_match("Centrifugal Pump", self.candidates, 75)
        assert isinstance(result["score"], int)


# ══════════════════════════════════════════════════════════════════════
# Tests: classify_match
# ══════════════════════════════════════════════════════════════════════
class TestClassifyMatch:
    def test_exact_match(self):
        status, conf = classify_match(100)
        assert status == "Exact Match"
        assert conf == "High"

    def test_exact_boundary(self):
        status, _ = classify_match(95)
        assert status == "Exact Match"

    def test_strong_match(self):
        status, conf = classify_match(90)
        assert status == "Strong Match"
        assert conf == "High"

    def test_partial_match(self):
        status, conf = classify_match(80)
        assert status == "Partial Match"
        assert conf == "Medium"

    def test_weak_match(self):
        status, conf = classify_match(60)
        assert status == "Weak Match"
        assert conf == "Low"

    def test_zero_score(self):
        status, conf = classify_match(0)
        assert status == "Weak Match"
        assert conf == "Low"


# ══════════════════════════════════════════════════════════════════════
# Tests: File readers
# ══════════════════════════════════════════════════════════════════════
class TestFileReaders:
    def test_read_aramco(self):
        data = {"Name": ["Centrifugal Pump", "Heat Exchanger"], "Id": ["AC-001", "AC-002"]}
        f = make_excel(data, sheet_name="ISM Functional Classes")
        records = read_aramco(f)
        assert len(records) == 2
        assert records[0]["name"] == "Centrifugal Pump"
        assert records[0]["source"] == "aramco"

    def test_read_aramco_fallback_sheet(self):
        """Falls back to sheet 0 if ISM Functional Classes doesn't exist."""
        data = {"Name": ["Pump"], "Id": ["1"]}
        f = make_excel(data, sheet_name="Sheet1")
        records = read_aramco(f)
        assert len(records) == 1

    def test_read_aramco_nan_filtering(self):
        """NaN names should be filtered out."""
        data = {"Name": ["Pump", float("nan"), ""], "Id": ["1", "2", "3"]}
        f = make_excel(data, sheet_name="ISM Functional Classes")
        records = read_aramco(f)
        assert len(records) == 1
        assert records[0]["name"] == "Pump"

    def test_read_cfihos(self):
        data = {"equipment class name": ["Pump", "Valve"], "CFIHOS unique id": ["CF-1", "CF-2"]}
        f = make_excel(data, sheet_name="Sheet1")
        records = read_cfihos(f)
        assert len(records) == 2
        assert records[0]["source"] == "cfihos"

    def test_read_kbr(self):
        data = {"Class Name (855)": ["Pump"], "Class Id": ["KBR-1"], "Discipline": ["Mechanical"]}
        f = make_excel(data, sheet_name="Sheet1")
        records = read_kbr(f)
        assert len(records) == 1
        assert records[0]["discipline"] == "Mechanical"
        assert records[0]["source"] == "kbr"

    def test_read_ltc_preferred_sheet(self):
        """Should prefer 'ISM Physical Classes' sheet."""
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            pd.DataFrame({"Name": ["Bad"], "Id": ["0"]}).to_excel(writer, sheet_name="ISM Functional Classes", index=False)
            pd.DataFrame({"Name": ["Good"], "Id": ["1"]}).to_excel(writer, sheet_name="ISM Physical Classes", index=False)
        buf.seek(0)
        records = read_ltc(buf)
        assert len(records) == 1
        assert records[0]["name"] == "Good"

    def test_read_ltc_strips_obsolete(self):
        """[OBSOLETE] prefix should be removed."""
        data = {"Name": ["[OBSOLETE] Old Pump"], "Id": ["1"]}
        buf = io.BytesIO()
        pd.DataFrame(data).to_excel(buf, sheet_name="ISM Physical Classes", index=False)
        buf.seek(0)
        records = read_ltc(buf)
        assert records[0]["name"] == "Old Pump"

    def test_read_sa_doc(self):
        data = {"Attribute": ["Weight", "Height"], "ID (CFIHOS_1.5)": ["1", "2"]}
        f = make_excel(data, sheet_name="SA_DOC_attributes")
        records = read_sa_doc(f)
        assert len(records) == 2
        assert records[0]["source"] == "sa_doc"


# ══════════════════════════════════════════════════════════════════════
# Tests: Attribute readers
# ══════════════════════════════════════════════════════════════════════
class TestAttributeReaders:
    def test_read_aramco_attributes(self):
        data = {
            "Class_Id": ["FC-001", "FC-001"],
            "Name": ["Weight", "Height"],
            "Id": ["A1", "A2"],
            "Discipline": ["Mechanical", "Mechanical"],
        }
        f = make_excel(data, sheet_name="ISM Functional Class Attributes")
        df = read_aramco_attributes(f)
        assert len(df) == 2
        assert "Discipline" in df.columns
        assert df["_source"].iloc[0] == "aramco"

    def test_read_aramco_attributes_backfill(self):
        """Name should be backfilled from Attribute_Desc when empty."""
        data = {
            "Class_Id": ["FC-001"],
            "Name": [None],
            "Lookup Att Desc": ["Weight Description"],
            "Id": ["A1"],
        }
        f = make_excel(data, sheet_name="ISM Functional Class Attributes")
        df = read_aramco_attributes(f)
        assert df["Name"].iloc[0] == "Weight Description"

    def test_read_aramco_attributes_wrong_sheet(self):
        """Should return empty DataFrame if sheet doesn't exist."""
        data = {"A": [1]}
        f = make_excel(data, sheet_name="WrongSheet")
        df = read_aramco_attributes(f)
        assert df.empty

    def test_read_ltc_attributes(self):
        data = {
            "Class_Id": ["PC-001"],
            "Name": ["Diameter"],
            "Id": ["A1"],
            "Discipline": ["Piping"],
        }
        f = make_excel(data, sheet_name="ISM Physical Class Attributes")
        df = read_ltc_attributes(f)
        assert len(df) == 1
        assert "Discipline" in df.columns
        assert df["_source"].iloc[0] == "ltc"

    def test_read_ltc_attributes_wrong_sheet(self):
        data = {"A": [1]}
        f = make_excel(data, sheet_name="WrongSheet")
        df = read_ltc_attributes(f)
        assert df.empty


# ══════════════════════════════════════════════════════════════════════
# Tests: Edge cases and integration
# ══════════════════════════════════════════════════════════════════════
class TestEdgeCases:
    def test_fuzzy_match_with_special_characters(self):
        candidates = [{"id": "1", "name": "Heat Exchanger (Shell & Tube)"}]
        result = fuzzy_match("Shell and Tube Heat Exchanger", candidates, 60)
        assert result is not None

    def test_fuzzy_match_compound_query(self):
        """Query with '/' should be expanded into parts."""
        candidates = [{"id": "1", "name": "Junction Box"}]
        result = fuzzy_match("Junction Box / Splice Case", candidates, 60)
        assert result is not None
        assert result["name"] == "Junction Box"

    def test_safe_str_pandas_nat(self):
        """pd.NaT should be handled."""
        assert _safe_str(pd.NaT) == ""

    def test_abbreviation_chaining(self):
        """Multiple abbreviations in sequence."""
        result = _normalize_abbreviations("CENTRIF PMP CTRL VLV")
        assert "centrifugal" in result
        assert "pump" in result
        assert "control" in result
        assert "valve" in result

    def test_tokenize_hyphenated(self):
        """Hyphenated words should be split by regex."""
        tokens = _tokenize("Air-Cooled")
        assert "air" in tokens
        assert "cooled" in tokens

    def test_classify_match_boundaries(self):
        """Test exact boundary values."""
        assert classify_match(95)[0] == "Exact Match"
        assert classify_match(94)[0] == "Strong Match"
        assert classify_match(85)[0] == "Strong Match"
        assert classify_match(84)[0] == "Partial Match"
        assert classify_match(75)[0] == "Partial Match"
        assert classify_match(74)[0] == "Weak Match"


# ══════════════════════════════════════════════════════════════════════
# Tests: Performance
# ══════════════════════════════════════════════════════════════════════
class TestPerformance:
    def test_fuzzy_match_performance(self):
        """500 candidates should complete in under 5 seconds."""
        candidates = [{"id": str(i), "name": f"Equipment Class Type {i}"} for i in range(500)]
        start = time.time()
        for _ in range(10):
            fuzzy_match("Equipment Class Type 250", candidates, 75)
        elapsed = time.time() - start
        assert elapsed < 5.0, f"10 fuzzy matches against 500 candidates took {elapsed:.2f}s"

    def test_expand_compound_performance(self):
        """1000 candidates with some compound names."""
        candidates = []
        for i in range(1000):
            name = f"Type {i}" if i % 3 != 0 else f"Type {i} / Variant {i}"
            candidates.append({"id": str(i), "name": name})
        start = time.time()
        names, idx_map = _expand_compound_names(candidates)
        elapsed = time.time() - start
        assert elapsed < 1.0
        assert len(names) > 1000  # compound names add more entries


# ══════════════════════════════════════════════════════════════════════
# Tests: Consistency between functions
# ══════════════════════════════════════════════════════════════════════
class TestConsistency:
    def test_fuzzy_match_uses_word_overlap(self):
        """Fuzzy match should reject matches without word overlap."""
        # "Valve" and "Pump" have no overlap — should NOT match even if fuzzy score is ok
        candidates = [{"id": "1", "name": "Pump"}]
        result = fuzzy_match("Valve", candidates, 0)
        assert result is None

    def test_abbreviation_and_overlap_work_together(self):
        """HX abbreviation should enable word overlap with Heat Exchanger."""
        candidates = [{"id": "1", "name": "Shell and Tube Heat Exchanger"}]
        result = fuzzy_match("Shell Tube HX", candidates, 50)
        assert result is not None

    def test_compound_and_abbreviation_together(self):
        """Compound name with abbreviation should match."""
        candidates = [{"id": "1", "name": "Heat Exchanger"}]
        result = fuzzy_match("HX / Condenser", candidates, 50)
        assert result is not None


# ══════════════════════════════════════════════════════════════════════
# Tests: Core harmonization pipeline
# ══════════════════════════════════════════════════════════════════════
class TestHarmonization:
    def test_single_master_harmonization(self):
        """Single master source harmonization."""
        files = {
            "aramco": {"records": [
                {"id": "AC-001", "name": "Centrifugal Pump", "source": "aramco"},
                {"id": "AC-002", "name": "Pressure Vessel", "source": "aramco"},
            ]},
            "kbr": {"records": [
                {"id": "KBR-1", "name": "Centrifugal Pump", "source": "kbr"},
            ]},
        }
        classes, reverse_gaps, logs = run_harmonization(["aramco"], files, 75)
        assert len(classes) == 2
        assert classes[0]["canonical_name"] == "Centrifugal Pump"
        assert "kbr" in classes[0]["matches"]
        assert "kbr" in classes[1]["gaps"]

    def test_dual_master_harmonization(self):
        """Dual master source harmonization."""
        files = {
            "aramco": {"records": [
                {"id": "AC-001", "name": "Centrifugal Pump", "source": "aramco"},
            ]},
            "cfihos": {"records": [
                {"id": "CF-001", "name": "Centrifugal Pump", "source": "cfihos"},
            ]},
            "kbr": {"records": [
                {"id": "KBR-1", "name": "Centrifugal Pump", "source": "kbr"},
            ]},
        }
        classes, reverse_gaps, logs = run_harmonization(["aramco", "cfihos"], files, 75)
        assert len(classes) == 1
        assert "aramco" in classes[0]["master_entries"]
        assert "cfihos" in classes[0]["master_entries"]
        assert "kbr" in classes[0]["matches"]

    def test_reverse_gaps_detected(self):
        """Non-master records that don't match any master entry are reverse gaps."""
        files = {
            "aramco": {"records": [
                {"id": "AC-001", "name": "Centrifugal Pump", "source": "aramco"},
            ]},
            "kbr": {"records": [
                {"id": "KBR-1", "name": "Centrifugal Pump", "source": "kbr"},
                {"id": "KBR-2", "name": "Cooling Tower", "source": "kbr"},
            ]},
        }
        classes, reverse_gaps, logs = run_harmonization(["aramco"], files, 75)
        assert "kbr" in reverse_gaps
        assert len(reverse_gaps["kbr"]) == 1
        assert reverse_gaps["kbr"][0]["name"] == "Cooling Tower"

    def test_demo_data_harmonization(self):
        """Full demo data should harmonize without errors."""
        demo = get_demo_data()
        files = {k: {"records": v} for k, v in demo.items()}
        classes, reverse_gaps, logs = run_harmonization(["aramco", "cfihos"], files, 75)
        assert len(classes) >= 8
        assert len(logs) >= 2


class TestBuildExportDf:
    def test_export_df_structure(self):
        """Export DataFrame should have correct columns."""
        demo = get_demo_data()
        files = {k: {"records": v} for k, v in demo.items()}
        classes, _, _ = run_harmonization(["aramco"], files, 75)
        df = build_export_df(classes, ["aramco"])
        assert "#" in df.columns
        assert "Gap In" in df.columns
        assert "Gap Count" in df.columns
        assert len(df) == len(classes)

    def test_export_df_gap_info(self):
        """Export should correctly label gaps."""
        files = {
            "aramco": {"records": [
                {"id": "AC-001", "name": "Centrifugal Pump", "source": "aramco"},
            ]},
            "kbr": {"records": []},
        }
        classes, _, _ = run_harmonization(["aramco"], files, 75)
        df = build_export_df(classes, ["aramco"])
        assert df.iloc[0]["Gap Count"] == 1
