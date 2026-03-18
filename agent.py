import pathlib
import shutil
import sys
import openpyxl
from typing import Literal, Optional

# Ensure UTF-8 output on Windows
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

from dotenv import load_dotenv
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage
from pydantic import BaseModel, Field
from langfuse.langchain import CallbackHandler
from langchain.chat_models import init_chat_model

from excel_read_helpers import excel_to_text_grid_full, excel_to_text_grid_values_only
from excel_write_helpers import write_cells

load_dotenv()

COMPANY_PROFILE = pathlib.Path("company_profile.md").read_text()

# llm = ChatOpenAI(model="gpt-5.4-2026-03-05", temperature=0)
llm = init_chat_model("google_genai:gemini-3.1-pro-preview", temperature=0)
langfuse_handler = CallbackHandler()


# ── Output schemas ─────────────────────────────────────────────────────────────

class ExtractedCell(BaseModel):
    """A single input cell identified in a German tender Excel form."""

    sheet: str = Field(description="The sheet name this cell belongs to")
    cell: str = Field(description="Cell coordinate, e.g. 'D5'")
    description: str = Field(
        description="Short German description of what this field is asking for")
    type: Literal["input", "checkbox"] = Field(
        description="'checkbox' for checkbox controls, 'input' for all other editable fields")
    current_value: str = Field(
        description="The cell's existing value, or empty string if the cell is blank")


class ExtractResponse(BaseModel):
    """All input cells found across every sheet of the workbook."""

    cells: list[ExtractedCell] = Field(
        description="Flat list of all identified input cells")


class AnalyzeResponse(BaseModel):
    """Result of the sheet-analysis pass."""

    cell_selection_rules: str = Field(
        description=(
            "Rules and criteria for identifying which cells are genuine input fields "
            "that must be filled in. Include any workbook-specific conventions, "
            "colour coding, numbering schemes, or explicit statements about which "
            "areas are editable vs. read-only."
        )
    )
    field_guidance: str = Field(
        description=(
            "All domain knowledge, definitions, legal references, examples, and "
            "contextual notes that describe *what content* belongs in the fields — "
            "everything that would help write an accurate description for each cell "
            "in the next step. Do not leave anything out."
        )
    )


class CellFix(BaseModel):
    """A single cell the validator flagged as needing correction."""

    cell: str = Field(description="Cell coordinate, e.g. 'D5'")
    reason: str = Field(
        description="Why this cell needs fixing, citing the exact value in [COMPANY DATA] that proves it")


class ValidationResult(BaseModel):
    """LLM verdict for a single sheet after filling."""

    accepted: bool = Field(
        description="True if the sheet is filled correctly")
    cells_to_fix: list[CellFix] = Field(
        default_factory=list,
        description="Cells that need fixing. Must be empty when accepted=True. Only include cells whose correct value is explicitly present in [COMPANY DATA].")


class MatchedCell(BaseModel):
    """A cell with a resolved fill value from company data."""

    sheet: str = Field(description="The sheet name this cell belongs to")
    cell: str = Field(description="Cell coordinate, e.g. 'D5'")
    value: Optional[str] = Field(
        description="Value to fill in from company data as a string, or null if unavailable. Only [CHECKBOX] cells use '1' to check, null to leave unchecked.")


class MatchResponse(BaseModel):
    """Fill instructions for every extracted cell."""

    cells: list[MatchedCell] = Field(
        description="Flat list of all cells with their resolved values")


# ── Step 0: Analyse sheets ────────────────────────────────────────────────────

def extract_instructions(file_path: str) -> AnalyzeResponse:
    """
    Single LLM call that reads all sheets and extracts:
    - cell_selection_rules: criteria for identifying which cells must be filled
    - field_guidance: domain knowledge / context needed to describe what goes in each cell
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet_names = [s.title for s in wb.worksheets if s.sheet_state == "visible"]
    wb.close()
    combined_block = ""
    for sheet_name in sheet_names:
        grid_text = excel_to_text_grid_values_only(file_path, sheet_name)
        combined_block += f"\n=== SHEET: {sheet_name} ===\n{grid_text}\n"

    system_prompt = """You are analysing a German tender/procurement Excel workbook (Vergabeformular).
You will receive the text content of all sheets, each delimited by === SHEET: <name> ===.

## Your task

Extract two distinct pieces of information and return them as separate fields:

### 1. cell_selection_rules
Rules and criteria that tell you *which* cells are genuine input fields requiring the applicant's data.
Include:
- Any explicit statements about which areas or columns are editable/fillable
- Colour-coding conventions (e.g. "yellow cells are mandatory, grey cells are optional")
- Numbering or structural conventions that mark input rows
- Any instructions about which sections to skip or leave blank
- Sheet-protection notes or legends

### 2. field_guidance
All domain knowledge and context that describes *what content* belongs in the fields.
Include:
- Definitions of terms, abbreviations, and German-law references
- Examples or sample values mentioned in the form
- Rules about required format, units, or length of answers
- Section headings and their explanations
- Any notes, footnotes, or legal references that clarify what a section is asking for

Do NOT invent content — only extract what is actually present in the workbook text.
If a field has nothing to report, return an empty string for that field."""

    human_prompt = f"ALL SHEETS OF THE WORKBOOK:\n{combined_block}"

    print(
        f"  [llm]    analysing {len(sheet_names)} sheets... ({len(combined_block):,} chars)")
    structured_llm = llm.with_structured_output(AnalyzeResponse)
    try:
        response: AnalyzeResponse = structured_llm.invoke(
            [SystemMessage(content=system_prompt),
             HumanMessage(content=human_prompt)],
            config={"callbacks": [langfuse_handler]},
        )
    except Exception as e:
        print(f"  [llm]    ERROR: {e}")
        raise

    print(
        f"  [analyse] cell_selection_rules={len(response.cell_selection_rules):,} chars, "
        f"field_guidance={len(response.field_guidance):,} chars"
    )
    return response


# ── Step 1: Extract ───────────────────────────────────────────────────────────
# add add sentence to look in colored cells (most likely yellow)
_EXTRACT_SYSTEM_PROMPT = """You are analysing a single sheet from a German tender/procurement Excel form (Vergabeformular).
You will receive:
- [CELL SELECTION RULES] — workbook-specific criteria for identifying which cells are genuine input fields
- [FIELD GUIDANCE] — domain knowledge, definitions, and context describing what content belongs in each field
- [SHEET GRID] — the actual cell data for the sheet

## Grid format

Each row is written as:  ROW_NUMBER: (COORD:value [tags...]), ...
- value is the cell text, or "" if empty
- [#RRGGBB]  — background colour
- [L]         — cell is locked (read-only); cells without [L] are editable
- [S]         — cell text is struck through (content is void/cancelled, do not fill)
- [CHECKBOX ✓"label"] — checked checkbox with optional label
- [CHECKBOX ○"label"] — unchecked checkbox with optional label

## How to identify REAL input fields

First, apply the [CELL SELECTION RULES] — they contain workbook-specific conventions
(colour coding, column conventions, explicit statements) that take priority.

For anything not covered by those rules, fall back to general judgement:
cells without [L] are potentially editable, but not every editable cell is an input field.
Ask yourself: "If I were filling this form by hand, would I write something here, or is this
just empty space, a formatting artifact, or a page border?"

Pay special attention to background colour as a strong signal for input fields:
- Yellow ([#FFFF00] or similar yellow tones) is the most common convention for mandatory input cells
- Other highlight colours (light green, light blue, orange) may also mark fillable fields
- If a colour appears repeatedly on clearly editable cells, treat it as an input-field marker even if not mentioned in [CELL SELECTION RULES]

## Sheets to SKIP entirely

If the sheet name or the first few visible cells of the sheet contain words such as
"alt", "old", "veraltet", "deprecated", "obsolet", "ungültig", "nicht verwenden", or similar
markers indicating the sheet is outdated or no longer in use — return an empty cells list.

## Cells to SKIP

Do NOT include the following — they are auto-generated or irrelevant to the applicant:
- Any cell tagged [S] — strikethrough text means the content is void or cancelled and must not be filled
- Date or timestamp fields (e.g. "Datum", "Date", cells containing or expecting a date/time value)
- Page numbers, print dates, auto-stamps
- Any cell whose sole purpose is to record when the document was printed or last modified

## How to generate descriptions

For each input field, build its description by combining all available context:
- **[FIELD GUIDANCE]** — definitions, legal references, and section explanations that describe what belongs in the field
- **Same row, earlier columns** — label cells to the left often name or number the field directly
- **Same column, earlier rows** — section headers or questions above give the topic this field belongs to
- **Empty label cells** — if the label columns in a row are empty, look upward to the nearest non-empty cell in those columns, and look left to the nearest non-empty cell in that row; forms often use merged cells that span multiple rows and columns, so a label shown on a previous row or an earlier column can apply to the current cell as well
- **Fill instructions** — if the form mentions how to fill the cell (e.g. "Bitte ankreuzen", "X eintragen", "ja oder nein", specific format like "TT.MM.JJJJ", or a list of allowed values), include that instruction verbatim in the description

Be specific, not generic. A form may have 5 different "Name" fields — your description must
disambiguate them. Include the section/question context and the specific data point being asked for,
so someone reading only the description knows exactly what to fill in without looking at the form.

## Output format

Return a flat JSON list under the key "cells". Group logically related multi-row fields
(e.g., 3 consecutive unlocked rows for the same question) into a single entry using the
top-left cell coordinate.

For `type`: use "checkbox" for any cell that contains a [CHECKBOX ...] tag; use "input" for
everything else. If you are unsure or the type is not explicitly indicated, default to "input".

For `current_value`: copy the cell's value exactly as it appears in the grid. Use "" if the
cell is empty. For checkboxes use "CHECKED" or "UNCHECKED" based on the ✓/○ symbol.

{
  "cells": [
    {
      "sheet": "Bewerbungsformblatt",
      "cell": "J9",
      "description": "2 - Name des Büros / Unternehmens",
      "type": "input",
      "current_value": "Mustermann GmbH"
    },
    {
      "sheet": "Bewerbungsformblatt",
      "cell": "J29",
      "description": "2.9 - Einheitliche Europäische Eigenerklärung (EEE) – ja oder nein",
      "type": "checkbox",
      "current_value": "UNCHECKED"
    },
    {
      "sheet": "Bewerbungsformblatt",
      "cell": "J38",
      "description": "3.1a - Anlage Nr. für Eintragung Inhaber/Führungskraft in Berufsregister",
      "type": "input",
      "current_value": ""
    }
  ]
}"""


def _extract_cells_for_sheet(
    sheet_name: str,
    grid_text: str,
    analysis: AnalyzeResponse,
    structured_llm,
    max_attempts: int = 10,
) -> list[dict]:
    """
    Consensus extraction for a single sheet.
    Repeats LLM calls until two attempts return identical (cell, type) sets,
    then returns that result. Falls back to >60%-frequency merge if no consensus.
    """
    from collections import Counter

    human_prompt = (
        f"[CELL SELECTION RULES]\n{analysis.cell_selection_rules}\n\n"
        f"[FIELD GUIDANCE]\n{analysis.field_guidance}\n\n"
        f"[SHEET GRID] (sheet: {sheet_name})\n{grid_text}"
    )

    attempts: list[list[dict]] = []

    for attempt in range(1, max_attempts + 1):
        print(
            f"  [consensus] '{sheet_name}' attempt {attempt}/{max_attempts} ({len(human_prompt):,} chars)")
        try:
            response: ExtractResponse = structured_llm.invoke(
                [SystemMessage(content=_EXTRACT_SYSTEM_PROMPT),
                 HumanMessage(content=human_prompt)],
                config={"callbacks": [langfuse_handler]},
            )
        except Exception as e:
            print(f"  [llm]    ERROR on '{sheet_name}': {e}")
            raise

        cells = [
            {
                "cell": item.cell,
                "sheet": sheet_name,
                "description": item.description,
                "type": item.type,
                "current_value": item.current_value,
            }
            for item in response.cells
        ]

        current_fp = frozenset((c["cell"], c["type"]) for c in cells)
        for i, prev in enumerate(attempts):
            if frozenset((c["cell"], c["type"]) for c in prev) == current_fp:
                print(
                    f"  [consensus] '{sheet_name}' — attempt {attempt} matches {i + 1}, {len(cells)} fields")
                return cells

        attempts.append(cells)

    # No two attempts agreed — keep cells appearing in >60% of attempts
    print(
        f"  [consensus] '{sheet_name}' — no consensus after {max_attempts}, merging by frequency")
    threshold = max_attempts * 0.5
    cell_counts: Counter = Counter(
        (c["cell"], c["type"]) for a in attempts for c in a)
    first_seen: dict[str, dict] = {}
    for a in attempts:
        for c in a:
            if c["cell"] not in first_seen:
                first_seen[c["cell"]] = c

    merged = [first_seen[coord]
              for (coord, _), count in cell_counts.items() if count > threshold]
    print(
        f"  [consensus] '{sheet_name}' — {len(merged)} cells kept (threshold >{threshold:.0f})")
    return merged


# ── Step 2: Match ─────────────────────────────────────────────────────────────

def match_cells(
    extracted: dict[str, list[dict]],
    field_guidance: str,
) -> dict[str, list[dict]]:
    sheets_block = ""
    for sheet_name, cells in extracted.items():
        sheets_block += f"\n=== SHEET: {sheet_name} ===\n"
        for item in cells:
            checkbox_note = f" [{item['type'].upper()}]" if item["type"] != "input" else ""
            current = f" [current: {item['current_value']}]" if item["current_value"] else ""
            sheets_block += f"  cell={item['cell']}{checkbox_note}{current}  — {item['description']}\n"

    system_prompt = """You are filling German tender Excel forms on behalf of a company.
You will receive:
- [FIELD GUIDANCE] — domain knowledge and context describing what each field is asking for
- [COMPANY DATA] — the company profile to draw fill values from
- [CELLS TO FILL] — the cells extracted from the workbook
Assign a fill value to each cell from the company data, or null if the data is unavailable.
Use [FIELD GUIDANCE] to understand exactly what each field expects so you can pick the most
appropriate value from the company data.

## Value rules

- Cells tagged [CHECKBOX]: use "1" to check, null to leave unchecked.
- Input cells asking for ja/nein or a selection marker: use "X".
- Input cells that specify a different exact marker in the form instructions: use that marker verbatim.
- All non-null values must be strings.
- Match the language of the form (German form → German values where appropriate).

Return a flat list of all cells under the key "cells":
{
  "cells": [
    {"sheet": "SheetName", "cell": "D5", "value": "Hansa Planungsgruppe GmbH"},
    {"sheet": "SheetName", "cell": "D7", "value": null}
  ]
}"""

    human_prompt = (
        f"[FIELD GUIDANCE]\n{field_guidance}\n\n"
        f"[COMPANY DATA]\n{COMPANY_PROFILE}\n\n"
        f"[CELLS TO FILL]\n{sheets_block}"
    )

    structured_llm = llm.with_structured_output(MatchResponse)
    response: MatchResponse = structured_llm.invoke(
        [SystemMessage(content=system_prompt),
         HumanMessage(content=human_prompt)],
        config={"callbacks": [langfuse_handler]},
    )

    lookup = {(m.sheet, m.cell): m for m in response.cells}
    for sheet_name, cells in extracted.items():
        for item in cells:
            matched = lookup.get((sheet_name, item["cell"]))
            item["value"] = matched.value if matched else None

    return extracted


# ── Step 3: Validate ─────────────────────────────────────────────────────────

_VALIDATE_SYSTEM_PROMPT = """You are a quality-control reviewer for German tender/procurement Excel forms.
You will receive:
- [SHEET NAME] — the sheet to review

## Sheets to skip
If [SHEET NAME] or the first visible cells of [FILLED SHEET GRID] indicate the sheet is outdated
(e.g. contains "alt", "old", "veraltet", "deprecated", "obsolet", "ungültig", "nicht verwenden"
or similar) — immediately return accepted=true with an empty cells_to_fix list.
- [FIELD GUIDANCE] — domain knowledge describing what each field expects
- [COMPANY DATA] — the authoritative company profile
- [FILLED SHEET GRID] — the current state of the sheet in the same grid format used during extraction

## Grid format reminder
Each row: ROW_NUMBER: (COORD:value [tags...]), ...
- [L] locked/read-only  [S] strikethrough  [CHECKBOX ✓/○]  [#RRGGBB] background colour

## Your task

Review only the sheet named in [SHEET NAME].

[COMPANY DATA] is the ONLY source of truth. Before flagging any cell you MUST confirm that the
correct value is explicitly present in [COMPANY DATA]. If it is not there, any blank or imperfect
value is acceptable — do NOT flag cells for data that simply does not exist in the profile.

Flag a cell ONLY if BOTH are true:
1. The cell is blank or contains a wrong value, AND
2. The correct value is explicitly and unambiguously present in [COMPANY DATA]

Always accept:
- Blank cells where the data does not exist in [COMPANY DATA]
- Cells requiring domain knowledge the profile does not cover
- Cells whose value is a reasonable interpretation of [COMPANY DATA]

Return:
- accepted: true if there are no fixable cells; false otherwise
- cells_to_fix: list of cells to correct — empty when accepted=True. For each entry include the
  cell coordinate and a reason that quotes the exact value from [COMPANY DATA] that should be used."""


def validate_sheet(
    file_path: str,
    sheet_name: str,
    field_guidance: str,
) -> ValidationResult:
    """Read the filled sheet with excel_to_text_grid_full and ask the LLM to identify fixable cells."""
    grid_text = excel_to_text_grid_full(file_path, sheet_name)
    human_prompt = (
        f"[SHEET NAME]: {sheet_name}\n\n"
        f"[FIELD GUIDANCE]\n{field_guidance}\n\n"
        f"[COMPANY DATA]\n{COMPANY_PROFILE}\n\n"
        f"[FILLED SHEET GRID]\n{grid_text}"
    )

    print(
        f"  [llm]    validating '{sheet_name}'... ({len(human_prompt):,} chars)")
    structured_llm = llm.with_structured_output(ValidationResult)
    try:
        result: ValidationResult = structured_llm.invoke(
            [SystemMessage(content=_VALIDATE_SYSTEM_PROMPT),
             HumanMessage(content=human_prompt)],
            config={"callbacks": [langfuse_handler]},
        )
    except Exception as e:
        print(f"  [llm]    ERROR validating '{sheet_name}': {e}")
        raise

    if result.accepted:
        print(f"  [validate] '{sheet_name}' — accepted")
    else:
        print(
            f"  [validate] '{sheet_name}' — {len(result.cells_to_fix)} cells to fix")
    return result


# ── Step 4: Fix flagged cells ─────────────────────────────────────────────────

_FIX_SYSTEM_PROMPT = """You are correcting specific cells in a German tender/procurement Excel form.
You will receive:
- [FIELD GUIDANCE] — domain knowledge describing what each field expects
- [COMPANY DATA] — the company profile as additional context
- [CELLS TO FIX] — cells flagged by the validator; each entry has the cell coordinate and an Issue
  that already states the exact correct value to use (quoted from the company profile)

## Primary rule
The Issue text for each cell is your primary instruction — it tells you exactly what value to write.
Use [COMPANY DATA] for clarification.

Apply the same value rules:
- [CHECKBOX] cells: "1" to check, null to leave unchecked.
- Input cells asking for ja/nein or a selection marker: use "X".
- Input cells that specify a different exact marker in the form instructions: use that marker verbatim.
- All non-null values must be strings

Return a flat list under the key "cells":
{
  "cells": [
    {"sheet": "SheetName", "cell": "D5", "value": "corrected value"},
    {"sheet": "SheetName", "cell": "D7", "value": null}
  ]
}"""


def fix_cells(
    sheet_name: str,
    cells_to_fix: list[CellFix],
    extracted_cells: list[dict],
    field_guidance: str,
) -> list[dict]:
    """Ask the LLM to provide corrected values for only the cells flagged by the validator."""
    desc_lookup = {c["cell"]: c for c in extracted_cells}

    cells_block = ""
    for fix in cells_to_fix:
        info = desc_lookup.get(fix.cell, {})
        type_note = " [CHECKBOX]" if info.get("type") == "checkbox" else ""
        cells_block += f"  {fix.cell}{type_note}\n  Issue: {fix.reason}\n\n"

    human_prompt = (
        f"[FIELD GUIDANCE]\n{field_guidance}\n\n"
        f"[COMPANY DATA]\n{COMPANY_PROFILE}\n\n"
        f"[CELLS TO FIX] (sheet: {sheet_name})\n{cells_block}"
    )

    print(
        f"  [llm]    fixing {len(cells_to_fix)} cells in '{sheet_name}'... ({len(human_prompt):,} chars)")
    structured_llm = llm.with_structured_output(MatchResponse)
    try:
        response: MatchResponse = structured_llm.invoke(
            [SystemMessage(content=_FIX_SYSTEM_PROMPT),
             HumanMessage(content=human_prompt)],
            config={"callbacks": [langfuse_handler]},
        )
    except Exception as e:
        print(f"  [llm]    ERROR fixing cells in '{sheet_name}': {e}")
        raise

    lookup = {m.cell: m for m in response.cells}
    result = []
    for fix in cells_to_fix:
        info = desc_lookup.get(fix.cell, {}).copy()
        info["cell"] = fix.cell
        info.setdefault("sheet", sheet_name)
        info.setdefault("type", "input")
        matched = lookup.get(fix.cell)
        info["value"] = matched.value if matched else None
        result.append(info)
    return result


# ── Orchestration ─────────────────────────────────────────────────────────────

def run(file_path: str) -> tuple[dict[str, list[dict]], dict[str, ValidationResult]]:
    src = pathlib.Path(file_path)
    out_path = pathlib.Path("tables_filled") / src.name
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Create output file upfront so per-sheet writes can update it in place
    shutil.copy2(str(src), str(out_path))
    print(f"\n-- {src.name} --")

    # Get sheet list and workbook-level instructions (once for the whole file)
    wb = openpyxl.load_workbook(str(src), data_only=True)
    sheet_names = [s.title for s in wb.worksheets if s.sheet_state == "visible"]
    wb.close()
    print(f"  [read]   {len(sheet_names)} sheets: {', '.join(sheet_names)}")

    analysis = extract_instructions(str(src))
    structured_extract_llm = llm.with_structured_output(ExtractResponse)

    all_cells: dict[str, list[dict]] = {}
    validation_results: dict[str, ValidationResult] = {}

    for sheet_name in sheet_names:
        print(f"\n  -- Sheet: '{sheet_name}' --")

        # Step 1 — extract input fields for this sheet (consensus loop)
        grid_text = excel_to_text_grid_full(str(src), sheet_name)
        sheet_cells = _extract_cells_for_sheet(
            sheet_name, grid_text, analysis, structured_extract_llm
        )

        if not sheet_cells:
            print(f"  [skip]   no input fields found in '{sheet_name}'")
            continue

        # Step 2 — initial match + write
        matched = match_cells({sheet_name: sheet_cells},
                              analysis.field_guidance)
        write_cells(str(out_path), {sheet_name: matched[sheet_name]})

        # Step 3+4 — validate → fix cycle (max 10 attempts)
        for attempt in range(1, 11):
            v_result = validate_sheet(
                str(out_path), sheet_name, analysis.field_guidance)

            if v_result.accepted or attempt == 10:
                all_cells[sheet_name] = matched[sheet_name]
                validation_results[sheet_name] = v_result
                if not v_result.accepted:
                    print(
                        f"  [validate] '{sheet_name}' — max attempts reached, moving on")
                break

            # Fix only the flagged cells and write them back
            fixed = fix_cells(sheet_name, v_result.cells_to_fix,
                              sheet_cells, analysis.field_guidance)
            write_cells(str(out_path), {sheet_name: fixed})
            # Merge fixes into matched so all_cells stays up to date
            fix_lookup = {c["cell"]: c for c in fixed}
            for item in matched[sheet_name]:
                if item["cell"] in fix_lookup:
                    item["value"] = fix_lookup[item["cell"]]["value"]

    return all_cells, validation_results
