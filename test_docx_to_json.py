#!/usr/bin/env python3
"""
Unit Tests for DOCX-to-JSON Conversion & Validation Tool (v3)
==============================================================
Fully generic tests that work with ANY structured test-case DOCX.
Uses sample/OIT4-Test_Cases.docx as a sample input for verification.

Run: pytest test_docx_to_json.py -v
"""

import json
import os
import shutil
import tempfile
from pathlib import Path
import pytest

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(BASE_DIR, "sample")
SAMPLE_DOCX = os.path.join(SAMPLE_DIR, "OIT4-Test_Cases.docx")

has_sample = os.path.exists(SAMPLE_DOCX)


# ── Fixtures ──

@pytest.fixture(scope="session")
def converted_data():
    """Run the parser on the sample DOCX and return parsed dict."""
    if not has_sample:
        pytest.skip("No sample DOCX found")
    from docx_to_json_tool import parse_test_cases_from_docx
    return parse_test_cases_from_docx(SAMPLE_DOCX)


@pytest.fixture(scope="session")
def all_cases(converted_data):
    cases = []
    for sec in converted_data["sections"]:
        cases.extend(sec["test_cases"])
    return cases


@pytest.fixture(scope="session")
def json_path(converted_data):
    """Write converted data to a temp file."""
    tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
    json.dump(converted_data, tmp, indent=2, ensure_ascii=False)
    tmp.close()
    yield tmp.name
    os.unlink(tmp.name)


@pytest.fixture(scope="session")
def jsonl_path(converted_data):
    """Write converted data to a temp JSONL file."""
    from docx_to_json_tool import generate_jsonl
    tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.jsonl', delete=False)
    tmp.close()
    generate_jsonl(converted_data, tmp.name)
    yield tmp.name
    os.unlink(tmp.name)


# ═══════════════════════════════════════════════════════════
# 1. JSON STRUCTURE (generic)
# ═══════════════════════════════════════════════════════════

class TestJsonStructure:

    def test_is_dict(self, converted_data):
        assert isinstance(converted_data, dict)

    def test_has_title(self, converted_data):
        assert "title" in converted_data
        assert converted_data["title"]

    def test_has_sections(self, converted_data):
        assert "sections" in converted_data
        assert isinstance(converted_data["sections"], list)
        assert len(converted_data["sections"]) > 0

    def test_has_coverage_list(self, converted_data):
        assert isinstance(converted_data.get("coverage"), list)

    def test_has_validation_summary(self, converted_data):
        assert isinstance(converted_data.get("validation_summary"), dict)


# ═══════════════════════════════════════════════════════════
# 2. TEST CASE COMPLETENESS (generic)
# ═══════════════════════════════════════════════════════════

class TestCompleteness:

    def test_has_test_cases(self, all_cases):
        assert len(all_cases) > 0

    def test_ids_unique(self, all_cases):
        ids = [tc["test_id"] for tc in all_cases]
        assert len(ids) == len(set(ids))

    def test_ids_continuous(self, all_cases):
        ids = sorted(tc["test_id"] for tc in all_cases)
        expected = list(range(ids[0], ids[-1] + 1))
        assert ids == expected

    def test_every_section_has_cases(self, converted_data):
        for sec in converted_data["sections"]:
            assert len(sec["test_cases"]) > 0, \
                f"Section {sec['section_number']} '{sec['section_title']}' is empty"

    def test_section_numbers_sequential(self, converted_data):
        nums = [s["section_number"] for s in converted_data["sections"]]
        assert nums == list(range(nums[0], nums[-1] + 1))


# ═══════════════════════════════════════════════════════════
# 3. PER-FIELD VALIDATION (generic)
# ═══════════════════════════════════════════════════════════

class TestFields:

    def test_every_case_has_prompt(self, all_cases):
        for tc in all_cases:
            assert tc["user_prompt"].strip(), f"Test {tc['test_id']}: empty prompt"

    def test_every_case_has_behaviors(self, all_cases):
        for tc in all_cases:
            assert isinstance(tc["expected_behavior"], list)
            assert len(tc["expected_behavior"]) > 0, f"Test {tc['test_id']}: no behaviors"
            for b in tc["expected_behavior"]:
                assert b.strip(), f"Test {tc['test_id']}: empty behavior string"

    def test_every_case_has_int_id(self, all_cases):
        for tc in all_cases:
            assert isinstance(tc["test_id"], int)

    def test_prompts_are_unique(self, all_cases):
        prompts = [tc["user_prompt"] for tc in all_cases]
        assert len(prompts) == len(set(prompts))

    def test_no_suspiciously_short_prompts(self, all_cases):
        for tc in all_cases:
            assert len(tc["user_prompt"]) >= 3, \
                f"Test {tc['test_id']}: prompt too short ({tc['user_prompt']})"


# ═══════════════════════════════════════════════════════════
# 4. SECTION STRUCTURE (generic)
# ═══════════════════════════════════════════════════════════

class TestSections:

    def test_section_has_number(self, converted_data):
        for sec in converted_data["sections"]:
            assert isinstance(sec["section_number"], int)

    def test_section_has_title(self, converted_data):
        for sec in converted_data["sections"]:
            assert isinstance(sec["section_title"], str)
            assert sec["section_title"].strip()

    def test_section_numbers_unique(self, converted_data):
        nums = [s["section_number"] for s in converted_data["sections"]]
        assert len(nums) == len(set(nums))

    def test_ids_sorted_within_sections(self, converted_data):
        for sec in converted_data["sections"]:
            ids = [tc["test_id"] for tc in sec["test_cases"]]
            assert ids == sorted(ids), \
                f"Section {sec['section_number']}: IDs not sorted"


# ═══════════════════════════════════════════════════════════
# 5. AI FOUNDRY COMPATIBILITY (generic)
# ═══════════════════════════════════════════════════════════

class TestAIFoundry:

    def test_flat_iteration(self, converted_data, all_cases):
        flat = []
        for sec in converted_data["sections"]:
            for tc in sec["test_cases"]:
                flat.append({"test_id": tc["test_id"],
                             "user_prompt": tc["user_prompt"],
                             "expected_behavior": tc["expected_behavior"]})
        assert len(flat) == len(all_cases)

    def test_serializable(self, all_cases):
        for tc in all_cases:
            reparsed = json.loads(json.dumps(tc))
            assert reparsed["user_prompt"] == tc["user_prompt"]

    def test_no_nulls(self, all_cases):
        for tc in all_cases:
            assert tc["user_prompt"] is not None
            assert tc["expected_behavior"] is not None


# ═══════════════════════════════════════════════════════════
# 6. JSONL OUTPUT (generic)
# ═══════════════════════════════════════════════════════════

class TestJsonlOutput:

    def test_jsonl_file_created(self, jsonl_path):
        assert os.path.exists(jsonl_path)

    def test_jsonl_line_count_matches(self, jsonl_path, all_cases):
        with open(jsonl_path, 'r') as f:
            lines = [l for l in f if l.strip()]
        assert len(lines) == len(all_cases)

    def test_jsonl_each_line_valid_json(self, jsonl_path):
        with open(jsonl_path, 'r') as f:
            for i, line in enumerate(f, 1):
                if not line.strip():
                    continue
                record = json.loads(line)
                assert "test_id" in record, f"Line {i}: missing test_id"
                assert "user_prompt" in record, f"Line {i}: missing user_prompt"
                assert "expected_behavior" in record, f"Line {i}: missing expected_behavior"
                assert "section" in record, f"Line {i}: missing section"

    def test_jsonl_prompts_match_json(self, jsonl_path, all_cases):
        with open(jsonl_path, 'r') as f:
            records = [json.loads(l) for l in f if l.strip()]
        for rec, tc in zip(records, all_cases):
            assert rec["user_prompt"] == tc["user_prompt"], \
                f"Test {rec['test_id']}: JSONL prompt mismatch"

    def test_jsonl_has_section_field(self, jsonl_path):
        with open(jsonl_path, 'r') as f:
            for line in f:
                if not line.strip():
                    continue
                rec = json.loads(line)
                assert rec["section"].strip(), f"Test {rec['test_id']}: empty section"


# ═══════════════════════════════════════════════════════════
# 7. EMBEDDED TEST SUITE (generic)
# ═══════════════════════════════════════════════════════════

class TestEmbeddedSuite:

    def test_suite_passes(self, json_path):
        from docx_to_json_tool import run_generic_test_suite
        report = run_generic_test_suite(json_path, SAMPLE_DOCX)
        assert report["overall_status"] == "PASS", \
            f"Failed: {report['failed_tests']}"

    def test_suite_has_enough_tests(self, json_path):
        from docx_to_json_tool import run_generic_test_suite
        report = run_generic_test_suite(json_path, SAMPLE_DOCX)
        assert report["total_tests"] >= 40


# ═══════════════════════════════════════════════════════════
# 8. DELTA-DIFF (generic)
# ═══════════════════════════════════════════════════════════

class TestDeltaDiff:

    @pytest.fixture(scope="class")
    def diff_report(self, json_path):
        from docx_to_json_tool import run_delta_diff
        return run_delta_diff(SAMPLE_DOCX, json_path)

    def test_passes(self, diff_report):
        assert diff_report["status"] == "PASS"

    def test_no_missing(self, diff_report):
        assert diff_report["missing_in_json"] == []

    def test_no_extra(self, diff_report):
        assert diff_report["extra_in_json"] == []

    def test_no_drift(self, diff_report):
        assert diff_report["drifted_prompts"] == []

    def test_count_match(self, diff_report):
        assert diff_report["count_match"]


# ═══════════════════════════════════════════════════════════
# 9. IDEMPOTENCY (parser determinism)
# ═══════════════════════════════════════════════════════════

class TestIdempotency:

    def test_two_runs_identical(self):
        """Parser must produce identical output on consecutive runs."""
        from docx_to_json_tool import parse_test_cases_from_docx, compare_json_outputs
        d1 = parse_test_cases_from_docx(SAMPLE_DOCX)
        d2 = parse_test_cases_from_docx(SAMPLE_DOCX)
        t1 = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
        t2 = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
        json.dump(d1, t1); t1.close()
        json.dump(d2, t2); t2.close()
        try:
            report = compare_json_outputs(t1.name, t2.name)
            assert report["status"] == "IDENTICAL"
        finally:
            os.unlink(t1.name)
            os.unlink(t2.name)


# ═══════════════════════════════════════════════════════════
# 10. OUTPUT FOLDER (generic)
# ═══════════════════════════════════════════════════════════

class TestOutputFolder:

    def test_folder_created(self):
        from docx_to_json_tool import create_output_folder
        folder = create_output_folder(SAMPLE_DOCX)
        try:
            assert os.path.isdir(folder)
            name = os.path.basename(folder)
            assert name.startswith("output_")
            parts = name.split("_")
            assert len(parts) >= 3
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_folder_has_parent_tag(self):
        from docx_to_json_tool import create_output_folder
        folder = create_output_folder(SAMPLE_DOCX)
        try:
            name = os.path.basename(folder)
            parent_tag = Path(SAMPLE_DOCX).resolve().parent.name.replace(" ", "-")[:40]
            assert parent_tag in name
        finally:
            shutil.rmtree(folder, ignore_errors=True)


# ═══════════════════════════════════════════════════════════
# 11. STRUCTURE VALIDATOR (generic)
# ═══════════════════════════════════════════════════════════

class TestStructureValidator:

    def test_passes(self, json_path):
        from docx_to_json_tool import validate_json_structure
        report = validate_json_structure(json_path)
        assert report["status"] == "PASS"
        assert report["errors"] == []
