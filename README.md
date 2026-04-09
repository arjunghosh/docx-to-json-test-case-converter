# DOCX-to-JSON Test Case Converter & Validator (v3)

A standalone Python CLI tool and library that converts structured Word (.docx) test-case documents into validated **JSON + JSONL** files suitable for **Azure AI Foundry** agent evaluation upload.

Built for the **Microsoft Co-Pilot Studio / Power BI / AI Foundry** project at Flexilytics.

---

## Features

- **DOCX Parsing** -- Extracts sections, test cases, user prompts, expected behaviors, pass/fail criteria, and preconditions from structured Word documents
- **Dual Output** -- Generates both `.json` (structured, hierarchical) and `.jsonl` (flat, one-test-per-line for AI Foundry upload)
- **Embedded QA Suite** -- Runs ~45 automated quality checks on every conversion (structure, completeness, field validation, AI Foundry compatibility, prompt quality, cross-reference consistency)
- **Delta-Diff Validator** -- Compares source DOCX against output JSON to detect prompt drift, missing/extra entries, and keyword loss
- **JSON Structure Validator** -- Validates schema compliance for AI Foundry compatibility
- **JSON-to-JSON Comparator** -- Detects drift/regression between two conversion runs
- **Idempotent Parser** -- Produces identical output on consecutive runs of the same input
- **Dual-Mode Design** -- Works as both a CLI tool (`python docx_to_json_tool.py convert ...`) and a Python library (`from docx_to_json_tool import convert_docx`)
- **Smart Text Normalization** -- Handles Unicode smart quotes, em-dashes, ellipses, and other special characters

---

## Project Structure

```
docx-to-json-test-case-converter/
  docx_to_json_tool.py        # Main tool (CLI + library API, ~1200 lines)
  test_docx_to_json.py         # Pytest suite (11 test classes, ~350 lines)
  README.md                    # This file
  .gitignore                   # Git ignore rules
  sample/
    OIT4-Test_Cases.docx       # Sample input for testing
```

---

## Architecture

The tool is organized into 10 sections within a single module:

| Section | Component | Description |
|---------|-----------|-------------|
| 1 | Output Folder Management | Creates timestamped output directories |
| 2 | DOCX Parser | State-machine parser for structured test-case documents |
| 3 | JSONL Generator | Flattens hierarchical data into AI Foundry format |
| 4 | Delta-Diff Validator | DOCX-vs-JSON drift detection with similarity scoring |
| 5 | JSON Structure Validator | Schema and field validation |
| 6 | Embedded Test Suite | ~45 generic QA checks (runs on any DOCX) |
| 7 | JSON-to-JSON Comparator | Regression detection between conversion runs |
| 8 | CLI Report Printers | Formatted terminal output for all reports |
| 9 | Library API | `convert_docx()` -- single-function pipeline for programmatic use |
| 10 | CLI Commands | `convert`, `validate`, `compare`, `full` subcommands |

---

## Requirements

- **Python** 3.10+
- **Dependencies:**

```bash
pip install python-docx pytest
```

---

## Usage

### 1. CLI Mode

```bash
# Convert a DOCX -> JSON + JSONL (with full QA pipeline)
python docx_to_json_tool.py convert sample/OIT4-Test_Cases.docx

# Absolute path
python docx_to_json_tool.py convert /Users/arjunghosh/Documents/my_tests.docx

# Path with spaces (use quotes)
python docx_to_json_tool.py convert "~/My Documents/test cases.docx"

# Validate an existing JSON against its source DOCX
python docx_to_json_tool.py validate source.docx output.json

# Compare two JSON outputs for drift/regression
python docx_to_json_tool.py compare old.json new.json
```

Every `convert` run automatically:
1. Creates a timestamped output folder: `output_<parent-tag>_<YYYYMMDD_HHMMSS>/`
2. Produces both `.json` (structured) and `.jsonl` (flat, one-test-per-line) files
3. Runs ~45 embedded QA tests across 11 categories
4. Runs a delta-diff between source DOCX and output JSON
5. Runs structure validation for AI Foundry compatibility
6. Writes a `validation_report.json` with all results
7. Prints a final CLI report with all file paths and quality score

### 2. Python Library Mode

```python
# OPTION A: Full pipeline (one function call)
from docx_to_json_tool import convert_docx

result = convert_docx("my_test_cases.docx")
print(result["status"])       # "PASS" or "FAIL"
print(result["json_path"])    # path to .json output
print(result["jsonl_path"])   # path to .jsonl output (for AI Foundry)

# With custom output directory
result = convert_docx("tests.docx", output_dir="./my_output")

# Silent mode (no print output)
result = convert_docx("tests.docx", quiet=True)

# Access parsed data directly
for section in result["data"]["sections"]:
    for tc in section["test_cases"]:
        print(tc["test_id"], tc["user_prompt"])
```

```python
# OPTION B: Individual functions (mix and match)
from docx_to_json_tool import (
    parse_test_cases_from_docx,   # DOCX -> dict
    generate_jsonl,                # dict -> .jsonl file
    validate_json_structure,       # check JSON schema
    run_delta_diff,                # compare DOCX vs JSON
    run_generic_test_suite,        # run ~45 QA checks
    compare_json_outputs,          # diff two JSONs
)

# Parse only
data = parse_test_cases_from_docx("my_tests.docx")

# Generate JSONL for AI Foundry upload
generate_jsonl(data, "output.jsonl")

# Validate
report = run_generic_test_suite("output.json", "my_tests.docx")
print(f"{report['passed']}/{report['total_tests']} checks passed")
```

### Library API Return Value

`convert_docx()` returns a dict with:

| Key | Type | Description |
|-----|------|-------------|
| `status` | `str` | `"PASS"` or `"FAIL"` |
| `json_path` | `str` | Path to output `.json` file |
| `jsonl_path` | `str` | Path to output `.jsonl` file |
| `report_path` | `str` | Path to `validation_report.json` |
| `output_dir` | `str` | Path to the output folder |
| `data` | `dict` | The parsed data (sections, test_cases, etc.) |
| `test_report` | `dict` | Embedded test suite results |
| `diff_report` | `dict` | Delta-diff results |
| `struct_report` | `dict` | Structure validation results |

---

## JSONL Format (for AI Foundry)

Each line in the `.jsonl` file is a self-contained test case:

```json
{"test_id": 1, "section": "Happy Path KPI Retrieval", "user_prompt": "What is total revenue for Q1 2025?", "expected_behavior": ["Classify as KPI Retrieval", "Generate DAX..."], "pass_criteria": "...", "fail_criteria": "..."}
```

Upload this file directly to Azure AI Foundry's evaluation framework.

---

## Running Tests

```bash
# Run the full pytest suite (11 test classes)
pytest test_docx_to_json.py -v
```

### Test Suite Coverage

| # | Test Class | What It Validates |
|---|------------|-------------------|
| 1 | `TestJsonStructure` | Top-level JSON structure (title, sections, coverage, validation_summary) |
| 2 | `TestCompleteness` | Test case count, ID uniqueness, ID continuity, sections not empty |
| 3 | `TestFields` | Every case has prompt, behaviors, integer ID; prompts unique and not too short |
| 4 | `TestSections` | Section numbers, titles, uniqueness, sorted IDs within sections |
| 5 | `TestAIFoundry` | Flat iteration, JSON serializability, no null values |
| 6 | `TestJsonlOutput` | JSONL file creation, line count, valid JSON per line, prompt matching |
| 7 | `TestEmbeddedSuite` | Embedded ~45-check QA suite passes with sufficient test count |
| 8 | `TestDeltaDiff` | DOCX-vs-JSON delta-diff passes with no missing/extra/drifted prompts |
| 9 | `TestIdempotency` | Two consecutive parser runs produce identical output |
| 10 | `TestOutputFolder` | Output folder creation with correct naming convention |
| 11 | `TestStructureValidator` | JSON structure validator passes with no errors |

---

## Expected DOCX Input Format

The tool expects Word documents structured as follows:

```
<Document Title>
<Designed For line (contains "->")>

It covers:
- Coverage item 1
- Coverage item 2

SECTION 1 --- Section Title

Test 1
User Prompt:
<prompt text>

Expected Behavior:
- bullet 1
- bullet 2

Pass if: <criteria>
Fail if: <criteria>

Test 2
...

SECTION 2 --- Another Section
...
```

### Supported Fields per Test Case

| Field | Required | Description |
|-------|----------|-------------|
| `test_id` | Yes | Integer test number |
| `user_prompt` | Yes | The user's input/question |
| `expected_behavior` | Yes | List of expected behavior bullets |
| `pass_criteria` | No | Pass condition string |
| `fail_criteria` | No | Fail condition string |
| `precondition` | No | Precondition text (parenthesized lines) |

---

## Output Structure

Each conversion run creates a timestamped folder:

```
output_sample_20250409_185000/
  OIT4-Test_Cases.json               # Structured JSON output
  OIT4-Test_Cases.jsonl              # Flat JSONL for AI Foundry
  OIT4-Test_Cases_validation_report.json  # Full validation report
```

---

## QA Pipeline (Embedded ~45 Checks)

The embedded test suite validates across 11 categories:

1. **JSON Structure** -- Valid JSON, required top-level fields
2. **Test Case Completeness** -- Case count, ID uniqueness, continuity
3. **Per-Field Validation** -- Non-empty prompts, behaviors, integer IDs
4. **Expected Behavior Content** -- No empty strings, all lists, unique prompts
5. **Special Fields** -- Preconditions, pass/fail criteria detection
6. **AI Foundry Compatibility** -- Flat iteration, serializability, no nulls
7. **Data Type Integrity** -- Coverage list, validation summary dict, section types
8. **Section Integrity** -- Unique/sequential section numbers, no empty sections
9. **Prompt Quality** -- Punctuation, length stats, question marks
10. **Delta-Diff (DOCX vs JSON)** -- Prompt count, drift, missing keywords
11. **Cross-Reference Consistency** -- Sorted IDs, position matching, totals

---

## License

Internal tool -- Flexilytics / Microsoft Co-Pilot Studio project.

---

## Author

Arjun Ghosh -- Flexilytics AI Projects
