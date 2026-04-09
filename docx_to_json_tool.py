#!/usr/bin/env python3
"""
DOCX-to-JSON Test Case Conversion & Validation Tool (v3)
=========================================================
Converts structured Word test-case documents into validated JSON + JSONL
suitable for Azure AI Foundry agent evaluation.

Works as BOTH a CLI tool and a Python library.

CLI Usage:
    # Convert DOCX -> JSON + JSONL (with full QA pipeline)
    python docx_to_json_tool.py convert /path/to/my_test_cases.docx

    # Absolute or relative paths both work
    python docx_to_json_tool.py convert ./sample/OIT4-Test_Cases.docx
    python docx_to_json_tool.py convert ~/Documents/test_cases.docx
    python docx_to_json_tool.py convert "C:/Users/me/My Tests/cases.docx"

    # Validate existing JSON against DOCX source
    python docx_to_json_tool.py validate source.docx output.json

    # Compare two JSON outputs for drift
    python docx_to_json_tool.py compare old.json new.json

Python Library Usage:
    from docx_to_json_tool import convert_docx

    # Full pipeline: returns dict with paths + validation results
    result = convert_docx("/path/to/my_test_cases.docx")
    print(result["json_path"])   # path to output JSON
    print(result["jsonl_path"])  # path to output JSONL
    print(result["status"])      # "PASS" or "FAIL"

    # Or use individual functions:
    from docx_to_json_tool import (
        parse_test_cases_from_docx,
        generate_jsonl,
        validate_json_structure,
        run_delta_diff,
        run_generic_test_suite,
        compare_json_outputs,
    )

    data = parse_test_cases_from_docx("my_tests.docx")
    # data is a dict with sections, test_cases, etc.
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

from docx import Document


# ─────────────────────────────────────────────────────────
# SECTION 1: OUTPUT FOLDER MANAGEMENT
# ─────────────────────────────────────────────────────────

def create_output_folder(docx_path: str) -> str:
    """
    Create a timestamped output folder next to the source DOCX.
    Format: output_<parent-folder-tag>_<YYYYMMDD_HHMMSS>/
    """
    docx_p = Path(docx_path).resolve()
    parent_dir = docx_p.parent
    parent_tag = parent_dir.name.replace(" ", "-")[:40]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"output_{parent_tag}_{timestamp}"
    output_dir = parent_dir / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)
    return str(output_dir)


# ─────────────────────────────────────────────────────────
# SECTION 2: DOCX PARSER
# ─────────────────────────────────────────────────────────

def normalize_text(text: str) -> str:
    """Normalize smart quotes and special Unicode to ASCII equivalents."""
    replacements = {
        "\u2018": "'", "\u2019": "'",
        "\u201c": '"', "\u201d": '"',
        "\u2026": "...",
        "\u2013": "-", "\u2014": "-",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def extract_paragraphs_from_docx(docx_path: str) -> list[str]:
    """Extract all non-empty paragraph texts from a DOCX file, normalized."""
    doc = Document(docx_path)
    return [normalize_text(p.text.strip()) for p in doc.paragraphs if p.text.strip()]


def parse_test_cases_from_docx(docx_path: str) -> dict:
    """
    Parse a structured test-case DOCX into a dict.

    Expected DOCX structure per test:
        SECTION N --- Section Title
        Test N
        User Prompt:
        <prompt text>
        Expected Behavior:
        - bullet 1
        - bullet 2
        Pass if: / Fail if: (optional)
    """
    paragraphs = extract_paragraphs_from_docx(docx_path)

    result = {
        "title": "",
        "designed_for": "",
        "coverage": [],
        "alignment": "",
        "sections": [],
        "validation_summary": {}
    }

    current_section = None
    current_test = None
    reading_mode = None
    i = 0

    # Footer/appendix patterns to stop parsing expected behavior
    FOOTER_PATTERNS = [
        "What This Validates", "If you want", "Just tell me which",
        "Excel Pass/Fail", "Live Booth Demo", "Failure Mode Debug",
        "Enterprise-Grade", "Client Demo Playbook",
    ]

    while i < len(paragraphs):
        line = paragraphs[i]

        # --- Title extraction (first line, heuristic) ---
        if i == 0 and not current_section:
            # Use first non-trivial line as title
            if len(line) > 10 and not line.lower().startswith("section"):
                result["title"] = line.strip()
                i += 1
                continue

        # --- Designed for (generic pattern) ---
        if not result["designed_for"] and "->" in line and not current_section:
            result["designed_for"] = line.strip()
            reading_mode = None
            i += 1
            continue

        # --- Coverage items ---
        if line.lower().startswith("it covers"):
            reading_mode = "coverage"
            i += 1
            continue

        if reading_mode == "coverage":
            if line.startswith(("\u2705", "- \u2705", "+")):
                clean = re.sub(r'^[\-\s]*\u2705\s*', '', line).strip()
                if clean:
                    result["coverage"].append(clean)
                i += 1
                continue
            elif "aligned" in line.lower():
                result["alignment"] = line.strip()
                reading_mode = None
                i += 1
                continue
            else:
                reading_mode = None

        # --- Alignment line ---
        if "aligned" in line.lower() and "100%" in line.lower() and not current_section:
            result["alignment"] = line.strip()
            i += 1
            continue

        # --- Section headers ---
        section_match = re.match(
            r'.*SECTION\s+(\d+)\s*[-\u2014\u2013]+\s*(.+)', line, re.IGNORECASE
        )
        if section_match:
            if current_section and current_test:
                current_section["test_cases"].append(current_test)
                current_test = None
            if current_section:
                result["sections"].append(current_section)
            current_section = {
                "section_number": int(section_match.group(1)),
                "section_title": section_match.group(2).strip(),
                "test_cases": []
            }
            reading_mode = None
            i += 1
            continue

        # --- Test ID ---
        test_match = re.match(r'^Test\s+(\d+)$', line, re.IGNORECASE)
        if test_match:
            if current_test and current_section:
                current_section["test_cases"].append(current_test)
            current_test = {
                "test_id": int(test_match.group(1)),
                "user_prompt": "",
                "expected_behavior": []
            }
            reading_mode = None
            i += 1
            continue

        # --- User Prompt marker ---
        if line.lower().startswith("user prompt"):
            reading_mode = "prompt"
            after_colon = re.sub(r'^User Prompt[:\s]*', '', line, flags=re.IGNORECASE).strip()
            if after_colon and current_test:
                current_test["user_prompt"] = after_colon
                reading_mode = None
            i += 1
            continue

        # --- Reading prompt text ---
        if reading_mode == "prompt" and current_test:
            current_test["user_prompt"] = line.strip()
            reading_mode = None
            i += 1
            continue

        # --- Expected Behavior marker ---
        if line.lower().startswith("expected behavior"):
            reading_mode = "expected"
            after_colon = re.sub(r'^Expected Behavior[:\s]*', '', line, flags=re.IGNORECASE).strip()
            if after_colon and current_test:
                for sub in after_colon.split('\n'):
                    sub = sub.strip()
                    if sub:
                        current_test["expected_behavior"].append(sub)
            i += 1
            continue

        # --- Reading expected behavior bullets ---
        if reading_mode == "expected" and current_test:
            if re.match(r'^Test\s+\d+$', line, re.IGNORECASE):
                continue
            if "SECTION" in line.upper():
                continue

            pass_match = re.match(r'^Pass if[:\s]*(.*)', line, re.IGNORECASE)
            fail_match = re.match(r'^Fail if[:\s]*(.*)', line, re.IGNORECASE)
            if pass_match:
                current_test["pass_criteria"] = pass_match.group(1).strip()
                reading_mode = None
                i += 1
                continue
            if fail_match:
                current_test["fail_criteria"] = fail_match.group(1).strip()
                reading_mode = None
                i += 1
                continue

            if line.startswith("(") and line.endswith(")"):
                current_test["precondition"] = line.strip("() ")
                i += 1
                continue

            # Stop on footer/appendix
            if any(stop in line for stop in FOOTER_PATTERNS):
                reading_mode = None
                i += 1
                continue

            clean = re.sub(r'^[\-\u2022\*]\s*', '', line).strip()
            if clean:
                sub_lines = [s.strip() for s in clean.split('\n') if s.strip()]
                current_test["expected_behavior"].extend(sub_lines)
            i += 1
            continue

        # --- Precondition ---
        if line.startswith("(") and line.endswith(")") and current_test:
            current_test["precondition"] = line.strip("() ")
            i += 1
            continue

        # --- Pass/Fail criteria outside expected mode ---
        if current_test:
            pass_match = re.match(r'^Pass if[:\s]*(.*)', line, re.IGNORECASE)
            fail_match = re.match(r'^Fail if[:\s]*(.*)', line, re.IGNORECASE)
            if pass_match:
                current_test["pass_criteria"] = pass_match.group(1).strip()
                i += 1
                continue
            if fail_match:
                current_test["fail_criteria"] = fail_match.group(1).strip()
                i += 1
                continue

        # --- Validation summary table ---
        if "What This Validates" in line:
            reading_mode = "summary"
            i += 1
            continue

        if reading_mode == "summary":
            if "\u2714" in line or "true" in line.lower():
                area = re.sub(r'[\u2714\u2705\s]+$', '', line).strip()
                if area:
                    result["validation_summary"][area] = True
            i += 1
            continue

        i += 1

    # Flush last test and section
    if current_test and current_section:
        current_section["test_cases"].append(current_test)
    if current_section:
        result["sections"].append(current_section)

    # Build validation_summary from section titles if not parsed from table
    if not result["validation_summary"] and result["sections"]:
        for sec in result["sections"]:
            result["validation_summary"][sec["section_title"]] = True

    return result


# ─────────────────────────────────────────────────────────
# SECTION 3: JSONL GENERATOR
# ─────────────────────────────────────────────────────────

def generate_jsonl(data: dict, jsonl_path: str):
    """
    Generate a JSONL file from the parsed data.
    Each line is one test case flattened for AI Foundry evaluation upload.
    Fields: test_id, section, user_prompt, expected_behavior, pass_criteria,
            fail_criteria, precondition (when present).
    """
    with open(jsonl_path, 'w', encoding='utf-8') as f:
        for section in data.get("sections", []):
            for tc in section.get("test_cases", []):
                record = {
                    "test_id": tc["test_id"],
                    "section": section["section_title"],
                    "user_prompt": tc["user_prompt"],
                    "expected_behavior": tc["expected_behavior"],
                }
                if "pass_criteria" in tc:
                    record["pass_criteria"] = tc["pass_criteria"]
                if "fail_criteria" in tc:
                    record["fail_criteria"] = tc["fail_criteria"]
                if "precondition" in tc:
                    record["precondition"] = tc["precondition"]
                f.write(json.dumps(record, ensure_ascii=False) + "\n")


# ─────────────────────────────────────────────────────────
# SECTION 4: DELTA-DIFF VALIDATOR
# ─────────────────────────────────────────────────────────

def extract_all_prompts_from_docx(docx_path: str) -> list[str]:
    """Extract all user prompts from the DOCX by pattern matching."""
    paragraphs = extract_paragraphs_from_docx(docx_path)
    prompts = []
    capture_next = False
    for p in paragraphs:
        if p.lower().startswith("user prompt"):
            after = re.sub(r'^User Prompt[:\s]*', '', p, flags=re.IGNORECASE).strip()
            if after:
                prompts.append(after)
            else:
                capture_next = True
            continue
        if capture_next:
            prompts.append(p.strip())
            capture_next = False
    return prompts


def extract_all_prompts_from_json(json_path: str) -> list[str]:
    """Extract all user_prompt values from the JSON file."""
    with open(json_path, 'r') as f:
        data = json.load(f)
    prompts = []
    for section in data.get("sections", []):
        for tc in section.get("test_cases", []):
            prompts.append(tc.get("user_prompt", ""))
    return prompts


def similarity(a: str, b: str) -> float:
    """Compute string similarity ratio between two strings."""
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()


def run_delta_diff(docx_path: str, json_path: str) -> dict:
    """Compare DOCX source against JSON output for drift detection."""
    docx_prompts = extract_all_prompts_from_docx(docx_path)
    json_prompts = extract_all_prompts_from_json(json_path)

    with open(json_path, 'r') as f:
        json_data = json.load(f)

    report = {
        "status": "PASS",
        "total_docx_prompts": len(docx_prompts),
        "total_json_prompts": len(json_prompts),
        "count_match": len(docx_prompts) == len(json_prompts),
        "missing_in_json": [],
        "extra_in_json": [],
        "drifted_prompts": [],
        "empty_expected_behaviors": [],
        "missing_test_ids": [],
        "section_count": len(json_data.get("sections", [])),
        "issues": []
    }

    for idx, dp in enumerate(docx_prompts):
        if idx < len(json_prompts):
            jp = json_prompts[idx]
            sim = similarity(dp, jp)
            if sim < 0.95:
                report["drifted_prompts"].append({
                    "index": idx + 1, "docx": dp, "json": jp,
                    "similarity": round(sim, 3)
                })
        else:
            report["missing_in_json"].append({"index": idx + 1, "prompt": dp})

    for idx in range(len(docx_prompts), len(json_prompts)):
        report["extra_in_json"].append({"index": idx + 1, "prompt": json_prompts[idx]})

    all_ids = []
    for section in json_data.get("sections", []):
        for tc in section.get("test_cases", []):
            all_ids.append(tc.get("test_id"))
            if not tc.get("expected_behavior"):
                report["empty_expected_behaviors"].append(tc.get("test_id"))

    expected_ids = list(range(1, max(all_ids) + 1)) if all_ids else []
    report["missing_test_ids"] = sorted(set(expected_ids) - set(all_ids))

    # Generic keyword spot-check: extract significant words from expected behaviors
    # in the DOCX and verify they appear in the JSON
    docx_paragraphs = extract_paragraphs_from_docx(docx_path)
    json_text = json.dumps(json_data).lower()
    in_expected = False
    missing_keywords = []
    for p in docx_paragraphs:
        if p.lower().startswith("expected behavior"):
            in_expected = True
            continue
        if in_expected and re.match(r'^(Test\s+\d+|.*SECTION\s+\d+)', p, re.IGNORECASE):
            in_expected = False
            continue
        if in_expected and any(stop in p for stop in [
            "What This Validates", "If you want", "Just tell me",
            "Excel Pass/Fail", "Live Booth", "Failure Mode",
            "Enterprise-Grade", "Client Demo"
        ]):
            in_expected = False
            continue
        if in_expected:
            words = re.findall(r'[A-Za-z]{5,}', p)
            significant = [w for w in words if w.lower() not in
                           {"this", "that", "with", "from", "have", "will", "been",
                            "does", "only", "also", "then", "them", "their",
                            "should", "would", "could", "about", "which", "there",
                            "these", "those", "after", "before", "because"}]
            for w in significant[:2]:
                if w.lower() not in json_text:
                    missing_keywords.append(f"'{w}' from: {p[:60]}")
                    break

    if missing_keywords:
        report["issues"].append({
            "type": "missing_keywords",
            "detail": f"{len(missing_keywords)} expected behavior keywords not found in JSON",
            "phrases": missing_keywords[:5]
        })

    has_issues = (
        report["missing_in_json"] or report["extra_in_json"] or
        report["drifted_prompts"] or report["empty_expected_behaviors"] or
        report["missing_test_ids"] or report["issues"]
    )
    report["status"] = "FAIL" if has_issues else "PASS"
    return report


# ─────────────────────────────────────────────────────────
# SECTION 5: JSON STRUCTURE VALIDATOR
# ─────────────────────────────────────────────────────────

def validate_json_structure(json_path: str) -> dict:
    """Validate JSON file structure for AI Foundry compatibility."""
    report = {
        "status": "PASS", "valid_json": False,
        "checks": [], "errors": [], "warnings": []
    }

    try:
        with open(json_path, 'r') as f:
            data = json.load(f)
        report["valid_json"] = True
        report["checks"].append("JSON syntax: PASS")
    except json.JSONDecodeError as e:
        report["valid_json"] = False
        report["errors"].append(f"Invalid JSON: {e}")
        report["status"] = "FAIL"
        return report

    for field in ["title", "sections"]:
        if field in data:
            report["checks"].append(f"Top-level '{field}': PRESENT")
        else:
            report["errors"].append(f"Missing top-level field: '{field}'")

    for field in ["designed_for", "coverage", "alignment", "validation_summary"]:
        if field in data:
            report["checks"].append(f"Top-level '{field}': PRESENT")
        else:
            report["warnings"].append(f"Optional field missing: '{field}'")

    sections = data.get("sections", [])
    if not isinstance(sections, list):
        report["errors"].append("'sections' must be an array")
    else:
        report["checks"].append(f"Sections count: {len(sections)}")

    all_test_ids = []
    total_tests = 0
    for sec_idx, section in enumerate(sections):
        if "section_number" not in section:
            report["errors"].append(f"Section {sec_idx}: missing 'section_number'")
        if "section_title" not in section:
            report["errors"].append(f"Section {sec_idx}: missing 'section_title'")
        if "test_cases" not in section:
            report["errors"].append(f"Section {sec_idx}: missing 'test_cases'")
            continue
        for tc in section.get("test_cases", []):
            total_tests += 1
            tid = tc.get("test_id")
            if "test_id" not in tc:
                report["errors"].append(f"Test case missing 'test_id' in section {sec_idx}")
            else:
                all_test_ids.append(tid)
            if "user_prompt" not in tc:
                report["errors"].append(f"Test {tid}: missing 'user_prompt'")
            elif not tc["user_prompt"].strip():
                report["errors"].append(f"Test {tid}: empty 'user_prompt'")
            if "expected_behavior" not in tc:
                report["errors"].append(f"Test {tid}: missing 'expected_behavior'")
            elif not isinstance(tc["expected_behavior"], list):
                report["errors"].append(f"Test {tid}: 'expected_behavior' must be array")
            elif len(tc["expected_behavior"]) == 0:
                report["errors"].append(f"Test {tid}: empty 'expected_behavior' array")

    report["checks"].append(f"Total test cases: {total_tests}")

    if len(all_test_ids) != len(set(all_test_ids)):
        dupes = [x for x in all_test_ids if all_test_ids.count(x) > 1]
        report["errors"].append(f"Duplicate test IDs: {set(dupes)}")
    else:
        report["checks"].append("Test ID uniqueness: PASS")

    if all_test_ids:
        expected = list(range(1, max(all_test_ids) + 1))
        missing = sorted(set(expected) - set(all_test_ids))
        if missing:
            report["warnings"].append(f"Non-continuous test IDs, missing: {missing}")
        else:
            report["checks"].append("Test ID continuity (1-N): PASS")

    vs = data.get("validation_summary", {})
    if vs:
        report["checks"].append(f"Validation summary areas: {len(vs)}")

    report["status"] = "FAIL" if report["errors"] else "PASS"
    return report


# ─────────────────────────────────────────────────────────
# SECTION 6: GENERIC EMBEDDED TEST SUITE
# ─────────────────────────────────────────────────────────

def run_generic_test_suite(json_path: str, docx_path: str = None) -> dict:
    """
    Run a comprehensive generic test suite on the converted JSON.
    These tests work on ANY structured test-case JSON.
    """
    with open(json_path, 'r') as f:
        data = json.load(f)

    all_cases = []
    for section in data.get("sections", []):
        all_cases.extend(section.get("test_cases", []))

    results = []

    def add(category: str, name: str, passed: bool, detail: str = ""):
        results.append({
            "category": category, "test": name,
            "status": "PASS" if passed else "FAIL", "detail": detail
        })

    # ── 1. JSON Structure ──
    cat = "JSON Structure"
    add(cat, "Valid JSON file", True)
    add(cat, "Top-level 'title' present", "title" in data)
    add(cat, "Top-level 'sections' present", "sections" in data)
    add(cat, "Title is non-empty", bool(data.get("title")))
    add(cat, "Sections is a list", isinstance(data.get("sections"), list))

    # ── 2. Test Case Completeness ──
    cat = "Test Case Completeness"
    n_cases = len(all_cases)
    n_sections = len(data.get("sections", []))
    add(cat, f"Test cases found: {n_cases}", n_cases > 0,
        f"{n_cases} test cases across {n_sections} sections")
    add(cat, f"Sections found: {n_sections}", n_sections > 0)

    all_ids = [tc.get("test_id") for tc in all_cases if "test_id" in tc]
    unique_ids = set(all_ids)
    add(cat, "Test IDs are unique", len(all_ids) == len(unique_ids),
        f"{len(all_ids)} IDs, {len(unique_ids)} unique")

    if all_ids:
        expected_seq = list(range(min(all_ids), max(all_ids) + 1))
        missing_ids = sorted(set(expected_seq) - unique_ids)
        add(cat, "Test IDs are continuous", len(missing_ids) == 0,
            f"Missing IDs: {missing_ids}" if missing_ids else "All sequential")

    # ── 3. Per-Field Validation ──
    cat = "Per-Field Validation"
    missing_prompts = [tc.get("test_id") for tc in all_cases
                       if not tc.get("user_prompt", "").strip()]
    add(cat, "Every case has non-empty user_prompt",
        len(missing_prompts) == 0,
        f"Missing in tests: {missing_prompts}" if missing_prompts else "All present")

    missing_behaviors = [tc.get("test_id") for tc in all_cases
                         if not tc.get("expected_behavior")]
    add(cat, "Every case has non-empty expected_behavior",
        len(missing_behaviors) == 0,
        f"Missing in tests: {missing_behaviors}" if missing_behaviors else "All present")

    non_int_ids = [tc.get("test_id") for tc in all_cases
                   if not isinstance(tc.get("test_id"), int)]
    add(cat, "Every test_id is an integer", len(non_int_ids) == 0)

    # ── 4. Expected Behavior Content Quality ──
    cat = "Expected Behavior Content"
    empty_behavior_items = []
    for tc in all_cases:
        for b in tc.get("expected_behavior", []):
            if not b or not b.strip():
                empty_behavior_items.append(tc.get("test_id"))
                break
    add(cat, "No empty strings in expected_behavior arrays",
        len(empty_behavior_items) == 0)

    behavior_not_list = [tc.get("test_id") for tc in all_cases
                         if not isinstance(tc.get("expected_behavior"), list)]
    add(cat, "expected_behavior is always a list", len(behavior_not_list) == 0)

    avg_behaviors = (sum(len(tc.get("expected_behavior", []))
                         for tc in all_cases) / max(len(all_cases), 1))
    add(cat, f"Average behaviors per test: {avg_behaviors:.1f}", avg_behaviors >= 1.0)

    all_prompts = [tc.get("user_prompt", "") for tc in all_cases]
    unique_prompts = set(all_prompts)
    add(cat, "All prompts are unique", len(all_prompts) == len(unique_prompts),
        f"{len(all_prompts)} total, {len(unique_prompts)} unique")

    short_prompts = [tc.get("test_id") for tc in all_cases
                     if len(tc.get("user_prompt", "")) < 3]
    add(cat, "No suspiciously short prompts (<3 chars)", len(short_prompts) == 0)

    # ── 5. Special Fields ──
    cat = "Special Fields"
    precond = [tc for tc in all_cases if "precondition" in tc]
    add(cat, f"Preconditions found: {len(precond)}", True,
        f"Tests: {[t['test_id'] for t in precond]}" if precond else "None")
    pass_crit = [tc for tc in all_cases if "pass_criteria" in tc]
    add(cat, f"Pass criteria found: {len(pass_crit)}", True)
    fail_crit = [tc for tc in all_cases if "fail_criteria" in tc]
    add(cat, f"Fail criteria found: {len(fail_crit)}", True)

    # ── 6. AI Foundry Compatibility ──
    cat = "AI Foundry Compatibility"
    flat = []
    for section in data.get("sections", []):
        for tc in section.get("test_cases", []):
            flat.append({"test_id": tc.get("test_id"),
                         "user_prompt": tc.get("user_prompt"),
                         "expected_behavior": tc.get("expected_behavior"),
                         "section": section.get("section_title")})
    add(cat, "Flat iteration possible", len(flat) == n_cases)

    serializable = True
    for tc in all_cases:
        try:
            json.dumps(tc)
        except (TypeError, ValueError):
            serializable = False
            break
    add(cat, "All test cases JSON-serializable", serializable)

    null_fields = [tc.get("test_id") for tc in all_cases
                   if tc.get("user_prompt") is None or tc.get("expected_behavior") is None]
    add(cat, "No null values in required fields", len(null_fields) == 0)

    # ── 7. Data Type Integrity ──
    cat = "Data Type Integrity"
    add(cat, "Coverage is a list", isinstance(data.get("coverage"), list))
    add(cat, "Validation summary is a dict", isinstance(data.get("validation_summary"), dict))
    sec_nums_valid = all(isinstance(s.get("section_number"), int) for s in data.get("sections", []))
    add(cat, "Section numbers are integers", sec_nums_valid)
    sec_titles_valid = all(isinstance(s.get("section_title"), str) and s["section_title"]
                           for s in data.get("sections", []))
    add(cat, "Section titles are non-empty strings", sec_titles_valid)

    # ── 8. Section Integrity ──
    cat = "Section Integrity"
    sec_numbers = [s.get("section_number") for s in data.get("sections", [])]
    add(cat, "Section numbers are unique", len(sec_numbers) == len(set(sec_numbers)))
    add(cat, "Section numbers are sequential",
        sec_numbers == list(range(min(sec_numbers), max(sec_numbers) + 1)) if sec_numbers else True)
    empty_sections = [s.get("section_number") for s in data.get("sections", []) if not s.get("test_cases")]
    add(cat, "No empty sections", len(empty_sections) == 0)

    # ── 9. Prompt Quality ──
    cat = "Prompt Quality"
    ends_with_punc = sum(1 for tc in all_cases if tc.get("user_prompt", "")[-1:] in ".?!")
    punc_pct = (ends_with_punc / max(n_cases, 1)) * 100
    add(cat, f"Prompts ending with punctuation: {punc_pct:.0f}%",
        punc_pct >= 50, f"{ends_with_punc}/{n_cases}")
    avg_len = sum(len(tc.get("user_prompt", "")) for tc in all_cases) / max(n_cases, 1)
    add(cat, f"Average prompt length: {avg_len:.0f} chars", avg_len >= 5)
    max_prompt = max((tc.get("user_prompt", "") for tc in all_cases), key=len, default="")
    min_prompt = min((tc.get("user_prompt", "") for tc in all_cases), key=len, default="")
    add(cat, f"Shortest prompt: {len(min_prompt)} chars", len(min_prompt) >= 3, f'"{min_prompt}"')
    add(cat, f"Longest prompt: {len(max_prompt)} chars", len(max_prompt) <= 500,
        f'"{max_prompt[:80]}..."' if len(max_prompt) > 80 else f'"{max_prompt}"')
    has_question = sum(1 for tc in all_cases if "?" in tc.get("user_prompt", ""))
    add(cat, f"Prompts with question marks: {has_question}/{n_cases}", True, "Informational only")

    # ── 10. Delta-Diff Integration ──
    if docx_path and os.path.exists(docx_path):
        cat = "Delta-Diff (DOCX vs JSON)"
        diff = run_delta_diff(docx_path, json_path)
        add(cat, "Delta-diff overall status", diff["status"] == "PASS")
        add(cat, f"Prompt count match (DOCX:{diff['total_docx_prompts']} vs JSON:{diff['total_json_prompts']})",
            diff["count_match"])
        add(cat, "No missing prompts", len(diff["missing_in_json"]) == 0)
        add(cat, "No extra prompts", len(diff["extra_in_json"]) == 0)
        add(cat, "No drifted prompts", len(diff["drifted_prompts"]) == 0)
        add(cat, "No empty expected behaviors", len(diff["empty_expected_behaviors"]) == 0)
        add(cat, "No missing test IDs in sequence", len(diff["missing_test_ids"]) == 0)

    # ── 11. Cross-Reference Consistency ──
    cat = "Cross-Reference Consistency"
    sorted_ok = all(
        [tc["test_id"] for tc in sec.get("test_cases", [])] == sorted(tc["test_id"] for tc in sec.get("test_cases", []))
        for sec in data.get("sections", [])
    )
    add(cat, "Test IDs sorted within each section", sorted_ok)
    positional_ok = all(
        sec.get("section_number") == idx + 1
        for idx, sec in enumerate(data.get("sections", []))
    )
    add(cat, "Section numbers match array position", positional_ok)
    section_total = sum(len(s.get("test_cases", [])) for s in data.get("sections", []))
    add(cat, "Section test totals match flat count", section_total == n_cases)

    # ── Summary ──
    passed = sum(1 for r in results if r["status"] == "PASS")
    failed = sum(1 for r in results if r["status"] == "FAIL")
    total = len(results)
    return {
        "total_tests": total, "passed": passed, "failed": failed,
        "pass_rate": f"{(passed / max(total, 1)) * 100:.1f}%",
        "overall_status": "PASS" if failed == 0 else "FAIL",
        "results": results,
        "failed_tests": [r for r in results if r["status"] == "FAIL"]
    }


# ─────────────────────────────────────────────────────────
# SECTION 7: JSON-TO-JSON COMPARATOR
# ─────────────────────────────────────────────────────────

def compare_json_outputs(old_path: str, new_path: str) -> dict:
    """Compare two JSON outputs to detect drift between runs."""
    with open(old_path, 'r') as f:
        old = json.load(f)
    with open(new_path, 'r') as f:
        new = json.load(f)

    report = {"status": "PASS", "metadata_diff": {}, "section_diff": [],
              "prompt_diffs": [], "behavior_diffs": [], "structural_diffs": []}

    for key in ["title", "designed_for", "alignment"]:
        if old.get(key) != new.get(key):
            report["metadata_diff"][key] = {"old": old.get(key), "new": new.get(key)}

    old_secs = len(old.get("sections", []))
    new_secs = len(new.get("sections", []))
    if old_secs != new_secs:
        report["structural_diffs"].append(f"Section count: old={old_secs}, new={new_secs}")

    def flat_map(data):
        m = {}
        for sec in data.get("sections", []):
            for tc in sec.get("test_cases", []):
                m[tc.get("test_id")] = {**tc, "_section": sec.get("section_title")}
        return m

    old_map, new_map = flat_map(old), flat_map(new)
    for tid in sorted(set(list(old_map) + list(new_map))):
        if tid not in old_map:
            report["structural_diffs"].append(f"Test {tid}: NEW")
            continue
        if tid not in new_map:
            report["structural_diffs"].append(f"Test {tid}: REMOVED")
            continue
        if old_map[tid].get("user_prompt") != new_map[tid].get("user_prompt"):
            report["prompt_diffs"].append({
                "test_id": tid,
                "old": old_map[tid].get("user_prompt"),
                "new": new_map[tid].get("user_prompt"),
                "similarity": round(similarity(
                    old_map[tid].get("user_prompt", ""),
                    new_map[tid].get("user_prompt", "")), 3)
            })
        if old_map[tid].get("expected_behavior") != new_map[tid].get("expected_behavior"):
            report["behavior_diffs"].append({
                "test_id": tid,
                "old_count": len(old_map[tid].get("expected_behavior", [])),
                "new_count": len(new_map[tid].get("expected_behavior", []))
            })

    has_issues = (report["metadata_diff"] or report["prompt_diffs"] or
                  report["behavior_diffs"] or report["structural_diffs"])
    report["status"] = "DRIFT DETECTED" if has_issues else "IDENTICAL"
    return report


# ─────────────────────────────────────────────────────────
# SECTION 8: CLI REPORT PRINTERS
# ─────────────────────────────────────────────────────────

def print_divider(char="=", width=70):
    print(char * width)

def print_test_suite_report(report: dict):
    print()
    print_divider()
    print(f"  QUALITY ASSURANCE TEST SUITE REPORT")
    print(f"  Status: {report['overall_status']}  |  "
          f"Passed: {report['passed']}/{report['total_tests']}  |  "
          f"Rate: {report['pass_rate']}")
    print_divider()
    current_cat = None
    for r in report["results"]:
        if r["category"] != current_cat:
            current_cat = r["category"]
            print(f"\n  [{current_cat}]")
        icon = "[ok]" if r["status"] == "PASS" else "[!!]"
        line = f"    {icon} {r['test']}"
        if r.get("detail"):
            line += f"  -- {r['detail']}"
        print(line)
    if report["failed_tests"]:
        print(f"\n  FAILURES ({report['failed']})")
        print_divider("-")
        for f in report["failed_tests"]:
            print(f"    [!!] [{f['category']}] {f['test']}")
            if f.get("detail"):
                print(f"         {f['detail']}")
    print()

def print_delta_diff_report(report: dict):
    print()
    print_divider()
    print(f"  DELTA-DIFF: DOCX vs JSON")
    print(f"  Status: {report['status']}")
    print_divider()
    print(f"  DOCX prompts: {report['total_docx_prompts']}")
    print(f"  JSON prompts: {report['total_json_prompts']}")
    print(f"  Count match:  {report['count_match']}")
    if report["drifted_prompts"]:
        print("\n  Drifted Prompts:")
        for d in report["drifted_prompts"]:
            print(f"    [!!] #{d['index']} (sim={d['similarity']}): DOCX='{d['docx']}' JSON='{d['json']}'")
    if report["missing_in_json"]:
        print("\n  Missing in JSON:")
        for m in report["missing_in_json"]:
            print(f"    [!!] #{m['index']}: {m['prompt']}")
    print()

def print_comparison_report(report: dict):
    print()
    print_divider()
    print(f"  JSON COMPARISON: OLD vs NEW")
    print(f"  Status: {report['status']}")
    print_divider()
    if report["structural_diffs"]:
        for d in report["structural_diffs"]:
            print(f"    [!!] {d}")
    if report["prompt_diffs"]:
        for d in report["prompt_diffs"]:
            print(f"    Prompt {d['test_id']}: sim={d['similarity']}")
    if report["behavior_diffs"]:
        for d in report["behavior_diffs"]:
            print(f"    Behavior {d['test_id']}: {d['old_count']} -> {d['new_count']} items")
    if report["status"] == "IDENTICAL":
        print("\n  No differences found. Outputs are identical.")
    print()

def print_final_summary(output_dir, json_path, jsonl_path, report_path, test_report, overall):
    print()
    print_divider("*")
    print(f"  FINAL RESULT: {overall}")
    print_divider("*")
    print(f"\n  Output Files:")
    print(f"    JSON Output:        {json_path}")
    print(f"    JSONL Output:       {jsonl_path}")
    print(f"    Validation Report:  {report_path}")
    print(f"    Output Folder:      {output_dir}")
    print(f"\n  Quality Score: {test_report['passed']}/{test_report['total_tests']} "
          f"tests passed ({test_report['pass_rate']})")
    if test_report["overall_status"] == "PASS":
        print(f"  Verdict: Ready for AI Foundry evaluation upload")
    else:
        print(f"  Verdict: {test_report['failed']} issue(s) require attention")
    print()
    print_divider("*")
    print()


# ─────────────────────────────────────────────────────────
# SECTION 9: LIBRARY API
# ─────────────────────────────────────────────────────────

def convert_docx(docx_path: str, output_dir: str = None, quiet: bool = False) -> dict:
    """
    Full conversion pipeline usable from Python code.

    Args:
        docx_path:  Path to the source DOCX file.
        output_dir: Optional output directory. If None, a timestamped
                    folder is created next to the DOCX.
        quiet:      If True, suppress all print output.

    Returns:
        dict with keys:
            status       - "PASS" or "FAIL"
            json_path    - path to output .json file
            jsonl_path   - path to output .jsonl file
            report_path  - path to validation_report.json
            output_dir   - path to the output folder
            data         - the parsed dict (sections, test_cases, ...)
            test_report  - embedded test suite results
            diff_report  - delta-diff results
            struct_report - structure validation results

    Example:
        from docx_to_json_tool import convert_docx

        result = convert_docx("my_tests.docx")
        print(result["status"])      # "PASS"
        print(result["json_path"])   # "/path/to/output/my_tests.json"
        print(result["jsonl_path"])  # "/path/to/output/my_tests.jsonl"

        # Access parsed data directly
        for section in result["data"]["sections"]:
            for tc in section["test_cases"]:
                print(tc["test_id"], tc["user_prompt"])
    """
    docx_path = str(Path(docx_path).resolve())
    if output_dir is None:
        output_dir = create_output_folder(docx_path)
    else:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        output_dir = str(Path(output_dir).resolve())

    docx_name = Path(docx_path).stem
    json_path = os.path.join(output_dir, f"{docx_name}.json")
    jsonl_path = os.path.join(output_dir, f"{docx_name}.jsonl")
    report_path = os.path.join(output_dir, f"{docx_name}_validation_report.json")

    # Convert
    data = parse_test_cases_from_docx(docx_path)
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    generate_jsonl(data, jsonl_path)

    # Validate
    test_report = run_generic_test_suite(json_path, docx_path)
    diff_report = run_delta_diff(docx_path, json_path)
    struct_report = validate_json_structure(json_path)

    status = "PASS" if (
        test_report["overall_status"] == "PASS" and
        diff_report["status"] == "PASS" and
        struct_report["status"] == "PASS"
    ) else "FAIL"

    # Write validation report
    full_report = {
        "timestamp": datetime.now().isoformat(),
        "source_docx": docx_path,
        "output_json": json_path,
        "output_jsonl": jsonl_path,
        "output_folder": output_dir,
        "overall_status": status,
        "test_suite": test_report,
        "delta_diff": diff_report,
        "structure_validation": struct_report
    }
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump(full_report, f, indent=2, ensure_ascii=False)

    if not quiet:
        total = sum(len(s["test_cases"]) for s in data["sections"])
        print(f"  Converted: {docx_path}")
        print(f"  Output:    {output_dir}")
        print(f"  Sections:  {len(data['sections'])}, Tests: {total}")
        print(f"  JSON:      {json_path}")
        print(f"  JSONL:     {jsonl_path}")
        print(f"  Status:    {status} ({test_report['passed']}/{test_report['total_tests']} checks)")

    return {
        "status": status,
        "json_path": json_path,
        "jsonl_path": jsonl_path,
        "report_path": report_path,
        "output_dir": output_dir,
        "data": data,
        "test_report": test_report,
        "diff_report": diff_report,
        "struct_report": struct_report,
    }


# ─────────────────────────────────────────────────────────
# SECTION 10: CLI COMMANDS
# ─────────────────────────────────────────────────────────

def cmd_convert(args):
    """Convert DOCX to JSON + JSONL with full test suite."""
    docx_path = str(Path(args.docx_file).resolve())
    output_dir = create_output_folder(docx_path)
    docx_name = Path(docx_path).stem
    json_path = os.path.join(output_dir, f"{docx_name}.json")
    jsonl_path = os.path.join(output_dir, f"{docx_name}.jsonl")
    report_path = os.path.join(output_dir, f"{docx_name}_validation_report.json")

    print(f"\n  Source:  {docx_path}")
    print(f"  Output:  {output_dir}")

    # Step 1: Convert
    print_divider("-")
    print("  STEP 1: Converting DOCX -> JSON + JSONL")
    data = parse_test_cases_from_docx(docx_path)
    total_tests = sum(len(s["test_cases"]) for s in data["sections"])
    print(f"  Parsed {len(data['sections'])} sections, {total_tests} test cases")

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"  JSON:  {json_path}")

    generate_jsonl(data, jsonl_path)
    print(f"  JSONL: {jsonl_path}")

    # Step 2: Run test suite
    print_divider("-")
    print("  STEP 2: Running Quality Assurance Tests")
    test_report = run_generic_test_suite(json_path, docx_path)
    print_test_suite_report(test_report)

    # Step 3: Delta-diff
    print("  STEP 3: Running Delta-Diff (DOCX vs JSON)")
    diff_report = run_delta_diff(docx_path, json_path)
    print_delta_diff_report(diff_report)

    # Step 4: Structure validation
    print("  STEP 4: Running Structure Validation")
    struct_report = validate_json_structure(json_path)
    print()
    print_divider()
    print(f"  JSON STRUCTURE VALIDATION: {struct_report['status']}")
    print_divider()
    for c in struct_report["checks"]:
        print(f"    [ok] {c}")
    for e in struct_report.get("errors", []):
        print(f"    [!!] {e}")
    for w in struct_report.get("warnings", []):
        print(f"    [--] {w}")
    print()

    # Step 5: Write validation report
    overall = "PASS" if (
        test_report["overall_status"] == "PASS" and
        diff_report["status"] == "PASS" and
        struct_report["status"] == "PASS"
    ) else "FAIL"

    full_report = {
        "timestamp": datetime.now().isoformat(),
        "source_docx": docx_path,
        "output_json": json_path,
        "output_jsonl": jsonl_path,
        "output_folder": output_dir,
        "overall_status": overall,
        "test_suite": test_report,
        "delta_diff": diff_report,
        "structure_validation": struct_report
    }
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump(full_report, f, indent=2, ensure_ascii=False)

    print_final_summary(output_dir, json_path, jsonl_path, report_path, test_report, overall)
    return overall, output_dir, json_path, jsonl_path


def cmd_validate(args):
    """Validate existing JSON against DOCX source."""
    docx_path = str(Path(args.docx_file).resolve())
    json_path = str(Path(args.json_file).resolve())
    print(f"\n  Validating: {json_path}")
    print(f"  Against:    {docx_path}")
    test_report = run_generic_test_suite(json_path, docx_path)
    print_test_suite_report(test_report)
    diff_report = run_delta_diff(docx_path, json_path)
    print_delta_diff_report(diff_report)
    struct_report = validate_json_structure(json_path)
    print_divider()
    print(f"  JSON STRUCTURE: {struct_report['status']}")
    print_divider()
    for c in struct_report["checks"]:
        print(f"    [ok] {c}")
    print()
    overall = "PASS" if (test_report["overall_status"] == "PASS" and
                         diff_report["status"] == "PASS" and
                         struct_report["status"] == "PASS") else "FAIL"
    print_divider("*")
    print(f"  OVERALL: {overall}  |  Tests: {test_report['passed']}/{test_report['total_tests']} ({test_report['pass_rate']})")
    print_divider("*")
    print()
    return overall

def cmd_compare(args):
    """Compare two JSON outputs."""
    old_path = str(Path(args.old_json).resolve())
    new_path = str(Path(args.new_json).resolve())
    print(f"\n  Comparing:\n    Old: {old_path}\n    New: {new_path}")
    report = compare_json_outputs(old_path, new_path)
    print_comparison_report(report)
    return report["status"]

def cmd_full(args):
    """Full pipeline: convert + test + validate + report."""
    return cmd_convert(args)


def main():
    parser = argparse.ArgumentParser(
        description="DOCX-to-JSON Test Case Converter & Validator (v3)",
        formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("convert", help="Convert DOCX to JSON+JSONL (timestamped output, runs tests)")
    p.add_argument("docx_file", help="Path to source DOCX file")

    p = sub.add_parser("validate", help="Validate existing JSON against DOCX source")
    p.add_argument("docx_file", help="Path to source DOCX file")
    p.add_argument("json_file", help="Path to JSON file to validate")

    p = sub.add_parser("compare", help="Compare two JSON outputs for drift")
    p.add_argument("old_json", help="Path to old/baseline JSON")
    p.add_argument("new_json", help="Path to new JSON to compare")

    p = sub.add_parser("full", help="Full pipeline: convert + test + validate + report")
    p.add_argument("docx_file", help="Path to source DOCX file")

    args = parser.parse_args()
    if args.command in ("convert", "full"):
        result = cmd_convert(args)
        sys.exit(0 if result[0] == "PASS" else 1)
    elif args.command == "validate":
        sys.exit(0 if cmd_validate(args) == "PASS" else 1)
    elif args.command == "compare":
        sys.exit(0 if cmd_compare(args) == "IDENTICAL" else 1)


if __name__ == "__main__":
    main()
