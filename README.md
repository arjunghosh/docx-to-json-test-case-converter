<p align="center">
  <h1 align="center">DOCX-to-JSON Test Case Converter & Validator</h1>
  <p align="center">
    <strong>Convert structured Word (.docx) test-case documents into validated JSON + JSONL for AI agent evaluation pipelines</strong>
  </p>
  <p align="center">
    <a href="#features">Features</a> &middot;
    <a href="#quick-start">Quick Start</a> &middot;
    <a href="#usage">Usage</a> &middot;
    <a href="#api-reference">API Reference</a> &middot;
    <a href="#faq">FAQ</a> &middot;
    <a href="#contributing">Contributing</a>
  </p>
  <p align="center">
    <img src="https://img.shields.io/badge/python-3.10%2B-blue?style=flat-square&logo=python&logoColor=white" alt="Python 3.10+">
    <img src="https://img.shields.io/badge/version-3.0.0-green?style=flat-square" alt="Version 3.0.0">
    <img src="https://img.shields.io/badge/license-MIT-orange?style=flat-square" alt="MIT License">
    <img src="https://img.shields.io/badge/QA%20checks-45%2B-brightgreen?style=flat-square" alt="45+ QA Checks">
    <img src="https://img.shields.io/badge/AI%20Foundry-compatible-purple?style=flat-square" alt="AI Foundry Compatible">
  </p>
  <p align="center">
    <a href="https://arjunghosh.github.io/docx-to-json-test-case-converter/">Landing Page</a> &middot;
    <a href="https://github.com/arjunghosh/docx-to-json-test-case-converter">GitHub Repository</a>
  </p>
</p>

---

## What Is This?

**docx-to-json-test-case-converter** is an open-source Python CLI tool and library that **converts structured Microsoft Word (.docx) test-case documents into validated JSON and JSONL files** for use with AI agent evaluation frameworks such as **Azure AI Foundry**, **Microsoft Co-Pilot Studio**, **Power BI AI agents**, and other LLM/AI testing pipelines.

It solves a common pain point: teams author test cases in Word, but AI evaluation platforms require structured data (JSON/JSONL). Manual conversion is slow, error-prone, and doesn't scale. This tool automates the entire workflow -- **parse, convert, validate, and report** -- in a single command, with **~45 built-in quality assurance checks** to ensure zero data loss.

**Also known as:** Word to JSON converter, DOCX parser for AI testing, test case extractor, AI Foundry JSONL generator, Co-Pilot Studio test converter, document-to-data pipeline.

---

## Who Is This For?

- **AI/ML Engineers** building evaluation datasets for AI agents, chatbots, or LLM applications
- **QA Engineers** managing test cases in Word documents who need structured output for automation
- **Microsoft Co-Pilot Studio developers** converting test scenarios for agent evaluation
- **Azure AI Foundry users** who need JSONL test data for evaluation framework upload
- **Power BI AI agent teams** validating BI agent responses against expected behaviors
- **Anyone** who has structured test cases in Word and needs them in JSON/JSONL format

---

## Features

- **DOCX Parsing** -- State-machine parser extracts sections, test cases, user prompts, expected behaviors, pass/fail criteria, and preconditions from structured Word documents
- **Dual Output** -- Generates both `.json` (structured, hierarchical) and `.jsonl` (flat, one-test-per-line) for direct AI Foundry upload
- **45+ QA Checks** -- Embedded test suite validates structure, completeness, field integrity, AI Foundry compatibility, prompt quality, and cross-reference consistency across 11 categories
- **Delta-Diff Validator** -- Compares source DOCX against output JSON to detect prompt drift, missing/extra entries, and keyword loss with similarity scoring
- **JSON Structure Validator** -- Validates schema compliance for AI evaluation platform compatibility
- **JSON-to-JSON Comparator** -- Detects drift and regression between two conversion runs
- **Idempotent Parser** -- Produces identical output on consecutive runs of the same input (deterministic)
- **Dual-Mode Design** -- Works as both a CLI tool (`python docx_to_json_tool.py convert ...`) and a Python library (`from docx_to_json_tool import convert_docx`)
- **Smart Text Normalization** -- Handles Unicode smart quotes, em-dashes, ellipses, and other special characters automatically
- **Zero Config** -- No configuration files needed; just point at a DOCX and run

---

## Quick Start

```bash
# 1. Clone the repository
git clone https://github.com/arjunghosh/docx-to-json-test-case-converter.git
cd docx-to-json-test-case-converter

# 2. Install dependencies
pip install python-docx pytest

# 3. Convert your test cases (full QA pipeline runs automatically)
python docx_to_json_tool.py convert your_test_cases.docx

# 4. Check the output
ls output_*/   # timestamped folder with .json, .jsonl, and validation report
```

---

## Project Structure

```
docx-to-json-test-case-converter/
  docx_to_json_tool.py        # Main tool -- CLI + library API (~1200 lines)
  test_docx_to_json.py         # Pytest suite -- 11 test classes (~350 lines)
  docs/
    index.html                 # GitHub Pages landing page
  README.md                    # This file
  LICENSE                      # MIT License
  .gitignore                   # Git ignore rules
  sample/                      # Sample DOCX files for testing
```

---

## Usage

### 1. CLI Mode

```bash
# Convert a DOCX -> JSON + JSONL (with full QA pipeline)
python docx_to_json_tool.py convert my_test_cases.docx

# Absolute or relative paths both work
python docx_to_json_tool.py convert ./sample/OIT4-Test_Cases.docx
python docx_to_json_tool.py convert /path/to/my_tests.docx

# Path with spaces (use quotes)
python docx_to_json_tool.py convert "~/My Documents/test cases.docx"

# Validate an existing JSON against its source DOCX
python docx_to_json_tool.py validate source.docx output.json

# Compare two JSON outputs for drift/regression
python docx_to_json_tool.py compare old.json new.json
```

**What happens on every `convert` run:**

1. Creates a timestamped output folder: `output_<parent-tag>_<YYYYMMDD_HHMMSS>/`
2. Produces both `.json` (structured) and `.jsonl` (flat, one-test-per-line) files
3. Runs ~45 embedded QA tests across 11 categories
4. Runs a delta-diff between source DOCX and output JSON
5. Runs structure validation for AI Foundry compatibility
6. Writes a `validation_report.json` with all results
7. Prints a final CLI report with file paths and quality score

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

# Silent mode (no print output -- ideal for library use)
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

---

## API Reference

### `convert_docx(docx_path, output_dir=None, quiet=False) -> dict`

Full conversion pipeline -- parse, convert, validate, report. Returns:

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

### Individual Functions

| Function | Description |
|----------|-------------|
| `parse_test_cases_from_docx(docx_path)` | Parse a DOCX file into a structured Python dict |
| `generate_jsonl(data, jsonl_path)` | Write flat JSONL from parsed data (AI Foundry format) |
| `validate_json_structure(json_path)` | Validate JSON schema compliance |
| `run_delta_diff(docx_path, json_path)` | Compare DOCX source vs JSON output for drift detection |
| `run_generic_test_suite(json_path, docx_path)` | Run ~45 QA checks across 11 categories |
| `compare_json_outputs(old_path, new_path)` | Diff two JSON outputs for regression detection |

---

## Architecture

Single-module design with 10 composable sections:

| # | Component | Description |
|---|-----------|-------------|
| 1 | Output Folder Management | Creates timestamped output directories |
| 2 | DOCX Parser | State-machine parser for structured test-case documents |
| 3 | JSONL Generator | Flattens hierarchical data into AI Foundry JSONL format |
| 4 | Delta-Diff Validator | DOCX-vs-JSON drift detection with similarity scoring |
| 5 | JSON Structure Validator | Schema and field validation for AI platform compatibility |
| 6 | Embedded Test Suite | ~45 generic QA checks (works on any structured DOCX) |
| 7 | JSON-to-JSON Comparator | Regression detection between conversion runs |
| 8 | CLI Report Printers | Formatted terminal output for all reports |
| 9 | Library API | `convert_docx()` -- single-function pipeline for programmatic use |
| 10 | CLI Commands | `convert`, `validate`, `compare`, `full` subcommands |

---

## JSONL Output Format (for Azure AI Foundry)

Each line in the `.jsonl` file is a self-contained test case ready for AI Foundry upload:

```json
{"test_id": 1, "section": "Happy Path KPI Retrieval", "user_prompt": "What is total revenue for Q1 2025?", "expected_behavior": ["Classify as KPI Retrieval", "Generate DAX..."], "pass_criteria": "...", "fail_criteria": "..."}
```

Upload this file directly to Azure AI Foundry's evaluation framework for AI agent testing.

---

## Running Tests

```bash
# Run the full pytest suite (11 test classes)
pytest test_docx_to_json.py -v
```

### Test Suite Coverage (11 Classes)

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

The tool parses Word documents structured as follows:

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
| `test_id` | Yes | Integer test number (parsed from `Test N`) |
| `user_prompt` | Yes | The user's input/question to the AI agent |
| `expected_behavior` | Yes | List of expected behavior bullet points |
| `pass_criteria` | No | Pass condition string (from `Pass if:`) |
| `fail_criteria` | No | Fail condition string (from `Fail if:`) |
| `precondition` | No | Precondition text (parenthesized lines) |

---

## Output Structure

Each conversion run creates a timestamped folder:

```
output_sample_20250409_185000/
  my_test_cases.json                    # Structured JSON output
  my_test_cases.jsonl                   # Flat JSONL for AI Foundry
  my_test_cases_validation_report.json  # Full validation report
```

---

## QA Pipeline (Embedded ~45 Checks)

The embedded test suite validates across 11 categories on every conversion:

1. **JSON Structure** -- Valid JSON, required top-level fields present
2. **Test Case Completeness** -- Case count, ID uniqueness, ID continuity
3. **Per-Field Validation** -- Non-empty prompts, behaviors, integer IDs
4. **Expected Behavior Content** -- No empty strings, all lists, unique prompts
5. **Special Fields** -- Preconditions, pass/fail criteria detection
6. **AI Foundry Compatibility** -- Flat iteration, JSON serializability, no nulls
7. **Data Type Integrity** -- Coverage list, validation summary dict, section types
8. **Section Integrity** -- Unique/sequential section numbers, no empty sections
9. **Prompt Quality** -- Punctuation, length stats, question marks
10. **Delta-Diff (DOCX vs JSON)** -- Prompt count match, drift detection, missing keywords
11. **Cross-Reference Consistency** -- Sorted IDs, position matching, section totals

---

## Use Cases

| Use Case | Description |
|----------|-------------|
| **Azure AI Foundry** | Convert test-case documents into JSONL for AI agent evaluation upload |
| **Microsoft Co-Pilot Studio** | Validate test coverage for Co-Pilot agent testing |
| **Power BI AI Agents** | Structured test case management for BI agent QA |
| **LLM/Chatbot Testing** | Generate structured test data from Word-authored test plans |
| **CI/CD Pipelines** | Integrate as a validation step in automated testing workflows |
| **Any AI Agent Testing** | Generic parser works with any structured test-case DOCX |

---

## FAQ

**How do I convert a Word document to JSON for AI testing?**
Install with `pip install python-docx`, then run `python docx_to_json_tool.py convert your_test_cases.docx`. The tool parses structured test cases and outputs validated JSON and JSONL files ready for Azure AI Foundry or any AI agent evaluation pipeline.

**What format does the DOCX file need to be in?**
The tool expects sections marked as `SECTION N --- Title`, tests as `Test N`, `User Prompt:` fields, `Expected Behavior:` bullet lists, and optional `Pass if:` / `Fail if:` criteria. See the [Expected DOCX Input Format](#expected-docx-input-format) section above.

**Can I use this as a Python library instead of CLI?**
Yes. Import `convert_docx` from `docx_to_json_tool` and call `result = convert_docx("my_tests.docx")`. It returns a dict with status, file paths, parsed data, and all validation reports. Individual functions are also importable.

**What is the JSONL format used for Azure AI Foundry?**
Each line is a self-contained JSON object with `test_id`, `section`, `user_prompt`, `expected_behavior` (array), and optional `pass_criteria`, `fail_criteria`, and `precondition` fields. Upload directly to AI Foundry's evaluation framework.

**Does this tool work with Microsoft Co-Pilot Studio test cases?**
Yes. Any structured test-case document following the expected DOCX format works, including those authored for Co-Pilot Studio, Power BI AI agents, and Azure AI Foundry.

**How many validation checks are run?**
Approximately 45 automated QA checks across 11 categories: JSON structure, completeness, field validation, expected behavior content, special fields, AI Foundry compatibility, data type integrity, section integrity, prompt quality, delta-diff, and cross-reference consistency.

**Is the parser deterministic / idempotent?**
Yes. Running the parser twice on the same DOCX produces byte-for-byte identical output. This is verified by the `TestIdempotency` test class.

**What Python version do I need?**
Python 3.10 or higher. The only runtime dependency is `python-docx`.

---

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## License

This project is licensed under the **MIT License** -- see the [LICENSE](LICENSE) file for details.

---

## Author

**Arjun Ghosh**
Founder, [Loyla.ai](https://www.loyla.ai/) -- Multi-Agentic AI Products Lab

Technology leader with 21+ years of experience in AI/ML, agentic AI, and enterprise software. Cornell-certified Chief AI Officer. Building open-source tools for the AI evaluation ecosystem.

- Website: [loyla.ai](https://www.loyla.ai/)
- GitHub: [@arjunghosh](https://github.com/arjunghosh)
- LinkedIn: [in/arjunghosh](https://linkedin.com/in/arjunghosh)
- X / Twitter: [@arjunghosh](https://twitter.com/arjunghosh)

---

<p align="center">
  <a href="https://www.loyla.ai/">
    <strong>Loyla.ai</strong>
  </a>
  -- Multi-Agentic AI Products Lab
  <br>
  <a href="https://arjunghosh.github.io/docx-to-json-test-case-converter/">Documentation</a> &middot;
  <a href="https://github.com/arjunghosh/docx-to-json-test-case-converter/issues">Report Bug</a> &middot;
  <a href="https://github.com/arjunghosh/docx-to-json-test-case-converter/issues">Request Feature</a>
</p>
