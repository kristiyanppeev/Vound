# Vound — Excel Tender Form Agent

Automatically fills German procurement / tender Excel workbooks (VgV, HOAI)
from a company profile, using an LLM for every decision.

---

## How it works

```
┌─────────────────────────────────────────────────────────────────────┐
│                         INPUT                                        │
│   tender.xlsx  (one or many)          company_profile.md            │
└────────────────┬────────────────────────────────┬───────────────────┘
                 │                                │
                 ▼                                ▼
     ┌───────────────────────┐      ┌─────────────────────────────┐
     │  1. classify_sheets   │      │  3. extract_profile_slices  │
     │                       │      │                             │
     │  LLM reads a preview  │      │  LLM copies the verbatim    │
     │  of every sheet and   │      │  parts of the profile that  │
     │  assigns a category:  │      │  belong to each category:   │
     │                       │      │                             │
     │  read_only            │      │  declaration   ──────────┐  │
     │  instructions         │      │  company_form  ──────┐   │  │
     │  declaration          │      │  ref_company   ───┐  │   │  │
     │  company_form         │      │  ref_personnel ─┐ │  │   │  │
     │  reference_company    │      │                 │ │  │   │  │
     │  reference_personnel  │      └─────────────────┼─┼──┼───┼──┘
     │  fee_offer            │                        │ │  │   │
     └───────────┬───────────┘                        │ │  │   │
                 │                                    │ │  │   │
                 ▼                                    │ │  │   │
     ┌───────────────────────┐                        │ │  │   │
     │  2. extract_workbook  │                        │ │  │   │
     │     _instructions     │                        │ │  │   │
     │                       │                        │ │  │   │
     │  LLM reads instruction│                        │ │  │   │
     │  sheets and pulls out:│                        │ │  │   │
     │  · cell_selection_    │                        │ │  │   │
     │    rules              │                        │ │  │   │
     │  · reference_rules    │                        │ │  │   │
     │  · personnel_         │                        │ │  │   │
     │    requirements       │                        │ │  │   │
     └───────────┬───────────┘                        │ │  │   │
                 │                                    │ │  │   │
        ┌────────┴────────────────────────────────┐   │ │  │   │
        │  4. assign_references                   │◄──┘ │  │   │
        │                                         │     │  │   │
        │  One LLM call sees ALL reference slots  │     │  │   │
        │  + ALL available references at once.    │     │  │   │
        │  Each reference assigned to exactly     │     │  │   │
        │  one slot — no duplicates across slots. │     │  │   │
        │  Multi-slot sheets (Referenz 1/2/3 side │     │  │   │
        │  by side) get one unique ref per column.│     │  │   │
        └────────┬────────────────────────────────┘     │  │   │
                 │                                       │  │   │
        ┌────────┴────────────────────────────────┐     │  │   │
        │  5. assign_personnel                    │◄────┘  │   │
        │                                         │        │   │
        │  One LLM call assigns team members to   │        │   │
        │  personnel slots. Each person used once │        │   │
        │  across different sheets. If a sheet has│        │   │
        │  multiple Referenz columns (one person  │        │   │
        │  with N project references), the same  │        │   │
        │  person is assigned to all sub-slots.  │        │   │
        └────────┬────────────────────────────────┘        │   │
                 │                                          │   │
                 ▼                                          │   │
┌──────────────────────────────────────────────────────────┼───┼──────┐
│                     FILL LOOP (per sheet)                 │   │      │
│                                                           │   │      │
│  reference_company sheets ◄───────────────────────────────┘   │      │
│  ┌──────────────────────────────────────────────────────────┐ │      │
│  │ filter_reference_data — extract just the assigned ref    │ │      │
│  │ read full grid (excel_to_text_grid_full)                 │ │      │
│  │ _consensus_fill ──► loop until 2 identical attempts      │ │      │
│  │                      or fallback: keep >50% frequency    │ │      │
│  │ write_cells                                              │ │      │
│  └──────────────────────────────────────────────────────────┘ │      │
│                                                                │      │
│  reference_personnel sheets ◄──────────────────────────────────┘      │
│  ┌──────────────────────────────────────────────────────────┐         │
│  │ filter_personnel_data — extract just the assigned person │         │
│  │ read full grid → _consensus_fill → write_cells           │         │
│  │ for multi-slot sheets: called once per Referenz column,  │         │
│  │ each call targets its column and uses the person's Nth   │         │
│  │ project reference (bio rows only filled on first call)   │         │
│  └──────────────────────────────────────────────────────────┘         │
│                                                                        │
│  company_form sheets ◄──────────────────────────────── company_form   │
│  ┌──────────────────────────────────────────────────────────┐         │
│  │ read full grid → _consensus_fill → write_cells           │         │
│  └──────────────────────────────────────────────────────────┘         │
│                                                                        │
│  declaration sheets ◄───────────────────────────────── declaration    │
│  ┌──────────────────────────────────────────────────────────┐         │
│  │ read full grid → _consensus_fill → write_cells           │         │
│  └──────────────────────────────────────────────────────────┘         │
│                                                                        │
│  fee_offer sheets ◄─────────────────────────────────── company_form   │
│  ┌──────────────────────────────────────────────────────────┐         │
│  │ read full grid → _consensus_fill → write_cells           │         │
│  │ (only identity fields filled — pricing left blank)       │         │
│  └──────────────────────────────────────────────────────────┘         │
└────────────────────────────────────────────────────────────────────────┘
                 │
                 ▼
        tables_filled/tender.xlsx
```

---

## Consensus fill loop

Every sheet is filled using `_consensus_fill`:

```
attempt 1  →  LLM returns cells A, B, C
attempt 2  →  LLM returns cells A, B, C  ✓ matches → done
              (if no match, keep going up to 10 attempts)
              (fallback: keep only cells seen in >50% of attempts)
```

This prevents one-off LLM inconsistencies from writing wrong cells.

---

## Sheet categories

| Category              | Description                                     | Data source                                |
| --------------------- | ----------------------------------------------- | ------------------------------------------ |
| `read_only`           | Scoring/evaluation sheets for the evaluator     | Skipped entirely                           |
| `instructions`        | Participation conditions and fill rules         | Extracted once, used as context            |
| `declaration`         | Legal declarations (Erklärung, Eigenerklärung)  | Company name, address, representative      |
| `company_form`        | Main application form                           | Full company master data + financials      |
| `reference_company`   | One or more project reference slots per sheet   | One unique reference per slot              |
| `reference_personnel` | One key person per sheet (with 1–N ref columns) | Person profile + Nth project reference     |
| `fee_offer`           | Hourly rates / HOAI pricing                     | Company identity only — pricing left blank |

---

## Key design decisions

**No duplicate references** — `assign_references` and `assign_personnel` see all
slots and all available data in a single call. Company references are each assigned
to exactly one slot. Personnel are each assigned to exactly one sheet; if a sheet has
multiple Referenz columns (one person with N project references side by side), the same
person is assigned to all sub-slots and each call targets the person's Nth project reference.

**No hallucinated data** — each fill call receives only the slice of company data
relevant to its sheet type. Pricing fields are explicitly left null because no
pricing data exists in the profile.

**One-shot fill** — cell identification and value assignment happen in a single LLM
call per attempt (no separate extract → match steps).

**Verbatim profile slices** — `extract_profile_slices` copies the profile word-for-word
into each slice; no summarising or rephrasing that could lose detail.

---

## Files

| File                     | Purpose                                                |
| ------------------------ | ------------------------------------------------------ |
| `agent_v2.py`            | Main agent — classification, slicing, assignment, fill |
| `excel_read_helpers.py`  | Grid extraction, checkbox reading, sheet utilities     |
| `excel_write_helpers.py` | Cell writing, VML checkbox updates                     |
| `company_profile.md`     | Company data used as the fill source                   |
| `tables/`                | Input Excel workbooks                                  |
| `tables_filled/`         | Output — filled copies of the workbooks                |

---

## Ideas for improvement

**Profile slice coverage**
If input fields are left unfilled, the likely cause is that `extract_profile_slices`
did not capture the relevant information from `company_profile.md` for that sheet
category — the slice fed to the fill LLM is missing the data it needs, so the LLM
correctly leaves those fields null. If this happens, the slice extraction prompt and
profile structure should be reviewed to ensure all required data is covered by the
appropriate slice.

**LLM judge step**
If the fill LLM is writing incorrect cells (wrong coordinates, misread labels,
hallucinated values), a lightweight judge step could be added after `_consensus_fill`
returns. The judge would receive the original sheet grid, the filled cells, and the
source data, and flag any cells that look inconsistent. Flagged cells could be
nulled out, retried, or surfaced as warnings in the output log. This adds one LLM
call per sheet but can significantly reduce silent errors.

---

## Usage

```bash
# Fill all workbooks in tables/
python agent_v2.py

# Fill a specific file
python agent_v2.py tables/my_tender.xlsx
```
