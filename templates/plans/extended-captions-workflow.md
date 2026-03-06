# Extended Captions (DOCX) — Workflow + Spec

This project’s report generator already supports **long captions stored in a separate DOCX table at the end of the document**, while keeping the HTML output unchanged (including the existing more/less UI).

Implementation is in the report generator at [`parse_extended_captions()`](../generate_report.py:749) and the merge step in [`main()`](../generate_report.py:3105).

---

## Goal

- Keep **SECTION 01 Posting Schedule** readable (no page-long captions inside the main schedule table).
- Store long captions in an **Extended Captions** section at the very end of the DOCX.
- Generator merges them by **exact Post Title match** and uses the long caption for HTML.

---

## DOCX Spec (what the generator recognizes)

### Placement

- The **Extended Captions** section should be the **last section** in the DOCX (end of document, after Account Interactions).

### Section header text (human-facing)

Recommended (the script will insert this):

- `SECTION 05`
- `Extended Captions`

### Table format (machine-facing)

Add **one table** with **two columns** and a header row:

| Column 1 header | Column 2 header |
|---|---|
| `Post Title` | `Extended Caption` |

Notes:

- Each **row** = one post with a long caption.
- **Post Title must match exactly** the Title text in SECTION 01.
- The Extended Caption cell can contain multiple paragraphs and line breaks; Word will store this as newline-separated text and the generator will preserve it.

---

## Notation for SECTION 01 (to keep the schedule clean)

When a post uses an extended caption, use a consistent marker in the **Caption** cell in SECTION 01.

Recommended options (pick one and use it consistently):

- `[EXT]`
- `[Extended]`
- leave the Caption cell empty

Important behavior detail: SECTION 01 captions that start with `[` are ignored by the schedule parser in [`parse_posts_table()`](../generate_report.py:559), which helps prevent placeholders from showing up.

---

## How to add a long caption (human steps)

1. In **SECTION 01 Posting Schedule**, keep the caption short or use the chosen marker.
2. Scroll to the end → **SECTION 05 Extended Captions**.
3. Add a new row to the table:
   - Column 1: paste the exact Post Title
   - Column 2: paste the full caption (can be very long)

Tip: In Word, adding a row is easiest by clicking in the last cell of the table and pressing **Tab**, or right-clicking → Insert → Rows Below.

---

## Script that will set up SECTION 05 automatically

We will provide a script at [`add_extended_captions.py`](../add_extended_captions.py) that:

- takes an input DOCX and produces a **new output DOCX**
- appends SECTION 05 + instruction text + the empty 2-column table
- is **idempotent** (if the section/table already exists, it won’t duplicate it)

Target usage:

```bash
python add_extended_captions.py "Content_Schedule_INPUT/SB_IG_February_Schedule.docx" "Content_Schedule_INPUT/SB_IG_February_Schedule_updated.docx"
```

