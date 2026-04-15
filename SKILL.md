---
name: office-skill-vn
description: >-
  Creates or edits Vietnamese Word (.docx): (A) administrative documents per Decree
   30/2020 — default workflow copies Mau_cong_van_ND30_tai_ve.docx (two layout tables)
  and fills cells; never overwrite that frozen file from scripts; do not rebuild header as plain paragraphs. (B) legal normative
  documents per Decree 78/2025/NĐ-CP (layout and typography per annex). When the user
  attaches an existing .docx, changes only requested text and preserves formatting. Use
  for Word, docx, cong van, to trinh,
  bao cao, luat, nghi dinh, thong tu, van ban quy pham phap luat, Nghi dinh 30,
  Nghi dinh 78, the thuc van ban, chinh sua noi dung, tra loi cong van, quoc huy,
  tieu ngu, chan ky, bang can chinh, Mau_cong_van_ND30_tai_ve,
  Mau_cong_van_ND30_ban_tu_script, fix_tra_loi_header_layout, can le, quoc hieu mot dong,
  gach ngang, Straight Connector, oxml_underline_after_motto.
---

# Word: administrative (Decree 30) and legal / normative (Decree 78)

## Purpose

When the user asks to **create** or **edit** a Word file, load this skill and pick the correct **branch**:

| Branch | Scope | Notes |
|--------|--------|------|
| **A** | Administrative / inter-agency style (reference **Decree 30/2020**) | **Default:** copy and fill **`Mau_cong_van_ND30_tai_ve.docx`** (tables + body). **Do not** overwrite that file with the generator. Optional generator output: **`Mau_cong_van_ND30_ban_tu_script.docx`**. |
| **B** | **Legal normative** instruments (**Decree 78/2025/NĐ-CP** annex) | **Different first page**: state motto **top-right**, issuing body **top-left**, symbol/title rules differ — **do not** use the Decree 30 script to build this cover |

- **New file (A)**: **Copy** `Mau_cong_van_ND30_tai_ve.docx` from this skill folder → **Save As** the output name → replace text **inside table cells and body paragraphs only**. See **Official Word template (branch A)** below.
- **New file (B)**: follow **Branch B** below; verify against the **official** text (Cong bao / national legal database) and [reference-van-ban-quy-pham-phap-luat.md](reference-van-ban-quy-pham-phap-luat.md).
- **Edit existing file**: change **only** the requested wording; **preserve** all formatting — see **In-place edits**; do **not** rebuild the whole document from script or plain paragraphs.

## Official Word template (branch A) — **Mau_cong_van_ND30_tai_ve.docx**

**Path (next to this `SKILL.md`):** `Mau_cong_van_ND30_tai_ve.docx`

**Frozen hand-tuned template — do not overwrite.** `Mau_cong_van_ND30_tai_ve.docx` is **authoritative**. **Never** save over it from scripts or agents. Generator twin: **`Mau_cong_van_ND30_ban_tu_script.docx`**. Workflow: **Copy** frozen file → **Save As** new name → edit cells only.

This file is the **canonical** layout for Decree-30-style công văn / trả lời công văn in this skill. Inspected structure:

| Structure | Role |
|-----------|------|
| **Table 0** — **3 rows × 2 columns** | **Left column:** issuing body line(s), then “Số / V/v …” line. **Right column:** “CỘNG HÒA…” + “Độc lập…” (state name + motto), then place + date. Row 3 may be spacer — **keep as in template**. |
| **Column widths (table 0)** | **Do not** use **equal** half-width columns (e.g. 50 % / 50 %) unless house style requires it. Equal narrow halves make **Quốc hiệu** **wrap** to a second line in Word. **Canonical in this skill:** left ~**2800** twips (~49 mm), right ~**6554** twips (~115 mm), row total **9354** twips — see `scripts/office_skill_cli.py` (`fix` subcommand). |
| **Alignment (table 0)** | **Row 1, left cell:** issuing body — **center**. **Row 1, right cell:** quốc hiệu + tiêu ngữ (+ horizontal line paragraph below) — **center**. **Row 2, left cell:** “Số…”, “V/v…” — **center** (frozen template). **Row 2, right cell:** địa danh + ngày — **right**. After edits, **restore** alignment if it was lost (common when assigning `paragraph.text` in python-docx). |
| **Quốc hiệu — one line** | **One paragraph, one line:** **CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM** (in hoa đậm theo thông lệ NĐ 30). **No** soft line break (`Shift+Enter`) inside the name. If Word still wraps visually, **widen the right column** first. |
| **Bold vs regular (table 0)** | **Left, row 1:** “TÊN CƠ QUAN CHỦ QUẢN” — **regular**. “TÊN CƠ QUAN BAN HÀNH” + “(dòng 2 nếu có)” — **bold**. **Right, row 1:** quốc hiệu + tiêu ngữ — **bold**. **Row 2, left:** “Số…”, “V/v…” — **regular**. **Row 2, right:** date — **italic**, not bold. Header rows: `w:spacing` after=0, line=288, lineRule=auto (frozen template). |
| **Gạch ngang dưới tiêu ngữ** | **Straight Connector** (`prst="line"`) in a **separate paragraph** under “Độc lập - …”, centered — **not** `w:pBdr` for a match to the frozen file. OOXML: `scripts/oxml_underline_after_motto.xml` (`parse_xml` in `scripts/office_skill_cli.py`). Word: **Insert → Shapes → Line**, or copy that paragraph from the frozen template. If missing after edits, restore via `ensure_underline_after_motto_cell` in `scripts/office_skill_cli.py`. |
| **Table 1 — chữ ký** | **Left:** “Nơi nhận” — **left**. **Right:** **one paragraph**, two **bold** runs: “KT. …” then **`w:br`** then “PHÓ …”; spacer lines; **bold** name — frozen template. |
| **Body** | Normal paragraphs after the tables (e.g. “Kính gửi:”, “Căn cứ …”, numbered items, closing). |
| **Table 1** — **1 row × 2 columns** | **Left:** “Nơi nhận: …”. **Right:** “KT. … / Phó …” + signer name block. |

**Correct workflow**

1. **Copy** `Mau_cong_van_ND30_tai_ve.docx` to the target path (or open → Save As). **Never** recreate the letterhead by typing Quốc hiệu / Tiêu ngữ / Cơ quan / Số / Ngày as **separate centered paragraphs** — that yields **0 tables** and **does not match** the template.
2. Edit **`document.tables[0]`** and **`document.tables[1]`** cells in python-docx, or edit the same cells in Word — preserve row/column count and merges.
3. Fill body paragraphs under the tables; keep **tab indents** (`\t`) where the template uses them.
4. **Template compliance rule:** unless the user explicitly requests global formatting changes, **do not** run “mass reformat” passes (e.g. forcing all paragraphs to justify, forcing font size across the entire document). Those operations commonly **break** the template’s run-level styling and spacing.

**Wrong output (reject this pattern)**

- A document where the “header” is **only** stacked **Normal** paragraphs (state name, motto, agency, number, date, subject) with **no tables** — this was the failure mode when the output diverged from the sample template. If the user supplied `Mau_cong_van_ND30_tai_ve.docx`, the result **must** still contain **two** layout tables like the template.

**When to use `scripts/office_skill_cli.py legacy`**

- Treat this as **legacy-only emergency path**. For ND30 reply letters, do **not** use it.
- Use only when both conditions are true:
  1. User explicitly rejects table-based template, and
  2. `Mau_cong_van_ND30_tai_ve.docx` is unavailable.
- State clearly that output will **not** match the official table-based template.

## In-place edits (content only, keep all formatting)

Use when the user attaches a `.docx` or gives a path and asks to revise, shorten, expand, rephrase, translate, etc.

### Mandatory rules

1. **Keep** margins, paper size, orientation, section breaks, **headers/footers** (page numbers, images, embedded signatures), **paragraph/character styles**, tabs, lists, **table structure**, frames, watermarks, hyperlinks — anything the user did **not** ask to change.
2. **Touch only** the passages the user treats as “content”. If scope is unclear (e.g. whether to change subject line or symbol), **ask once** before editing.
3. **Do not** replace the whole file with output from the template script — that **drops** bespoke formatting.
4. **Do not** “select all → paste” or rebuild in a fresh `Document()` then copy the full text (loses run-level formatting). Edit **in place** (`Document(path)`), save over the file or to `_sua.docx` if the user wants to keep the original.

### python-docx hints

- Open `Document(path)`.
- Locate paragraphs by existing text, headings, or indices the user gives.
- If a paragraph has **one** uniform run, `paragraph.text = "..."` may be OK; **always** re-apply **`paragraph.alignment`** (and cell-appropriate `WD_ALIGN_PARAGRAPH`) after `paragraph.text = ...` on **table0 / table 1** cells — python-docx often drops `w:jc`, which breaks center/left/right on quốc hiệu, ngày, and chữ ký.
- If a paragraph has **multiple runs** (mixed bold/italic), prefer editing **`run.text`** per run.
- Tables: edit only allowed cells; do not delete rows/columns/merges unless requested.
- Headers/footers: **leave unchanged** unless the user asks to change them.
- After saving, if risk of layout loss, suggest opening Word and comparing, or saving under a new name.

### If python-docx is not enough

Macros, content controls, or very custom OOXML: prefer Word UI or another approved tool and **state limitations** instead of destroying layout.

## Table layout for header and signatures (alignment)

Use this when the user works on **reply letters / inter-agency letters** (Vietnamese: tra loi cong van) where the **national emblem (quoc huy)**, **motto (tieu ngu)**, and **signature block (chan ky)** must stay aligned. **Do not** rely only on spaces or tabs; use **borderless Word tables** so columns stay fixed when text changes.

**Concrete reference:** `Mau_cong_van_ND30_tai_ve.docx` already implements this (table 0 + table 1). **Prefer copying that file** instead of improvising a new table layout.

### Header block (emblem + motto)

1. Insert a **table** at the top of the first page (or in the header area if the user’s template uses header — follow their file).
2. Set **no borders** (Table Design → **No border**, or Borders → **No Border**) so the grid is invisible when printed.
3. **Typical patterns** (pick what matches the user’s sample `.docx`):
   - **Two columns, one row**: left cell — **Quoc huy** (picture, fixed width ~2–2.5 cm or as in template); right cell — **Quoc hieu** (uppercase, bold), paragraph(s) **right-aligned** or **centered** in that cell per house style.
   - **Second row**, **merge cells** across full width: **Tieu ngu** centered, bold; optional horizontal line under the motto (shape or bottom border on paragraph only — not full table border).
4. Set **fixed column widths** (Table Properties → Column → Preferred width in cm or %) so layout does not jump when editing text.
5. **Vertical alignment**: **Center** for the cell that contains the emblem image; **Top** or **Center** for text cells as needed.
6. When **editing** an existing file: change **text inside cells only**; **do not** delete the table or merge/split cells unless the user asks.

### Signature block (chan ky)

1. Use a **borderless table** with **two or three columns** for parallel blocks (e.g. left: “KT. …”, right: “TM. …”, or multiple agencies).
2. Each column: **Chuc vu** (job title, uppercase bold) → blank lines for signature scan → **Ho va ten** (bold); align **center** within the column.
3. Fix **column widths** (often equal %, e.g. 50% / 50%) so left and right blocks stay balanced.
4. **Noi nhan** (if present) can sit in a **separate full-width row below** the signature table, or in its own borderless one-column table — match the user’s template.
5. When **editing**: only replace wording in the relevant **cells**; preserve row/column count and merges.

### Distribution block (Noi nhan) — strict formatting

Apply these defaults unless the user template explicitly differs:

1. Label line `Noi nhan:`: **Times New Roman**, **12 pt**, **bold + italic** (no underline), **left-aligned** — apply to the **first non-empty** paragraph in the cell (skip leading empty paragraphs).
2. Recipient lines (`- ...;`, `- Luu: ...`) and all lines below `Noi nhan:`: **Times New Roman**, **11 pt**, **regular (not bold/italic)**, **left-aligned**.
3. Paragraph spacing in the `Noi nhan` cell: **Before = 0 pt, After = 0 pt** for every line.
4. Keep one recipient per line and preserve punctuation (`;` on list lines, final punctuation per template).

### python-docx

- Treat layout tables like any table: edit via `table.rows[r].cells[c]` and paragraphs inside each cell.
- To **add** a similar structure in code: create `Table(...)` with desired rows/columns, set cell widths via `tcW`, remove borders via `tblBorders` if needed — or recommend **Word UI** for one-off templates.
- **Never** flatten a layout table to plain paragraphs when the user only asked for content changes.

### Script `office_skill_cli.py legacy` (legacy, blocked by default)

- This script builds a **simple centered stack** without the 2-table ND30 layout.
- It is now **blocked by default** and requires explicit `--allow-legacy-stack`.
- For ND30-compliant outputs, use `Mau_cong_van_ND30_tai_ve.docx` + `python scripts/office_skill_cli.py rebuild ...`.

## Helper script — legacy centered stack (do not use for ND30 template output)

- **Default for new branch A files:** use **`Mau_cong_van_ND30_tai_ve.docx`** (see above).
- The legacy script below is for exceptional fallback only and needs explicit opt-in flag.
- Path: `scripts/office_skill_cli.py` (inside this skill folder), subcommand `legacy`.
- Install: `pip install python-docx`

**PowerShell** (line continuation with `` ` ``):

```powershell
python scripts/office_skill_cli.py legacy --allow-legacy-stack --output "van_ban.docx" `
  --co-quan "TEN CO QUAN BAN HANH" `
  --so-ky-hieu "12/CV-ABC" `
  --dia-danh-ngay "Ha Noi, ngay 15 thang 4 nam 2026" `
  --trich-yeu "V/v …" `
  --body-pt 13 `
  --line-exact-pt 16
```

**CMD** (`^`):

```bat
python scripts/office_skill_cli.py legacy --allow-legacy-stack --output "van_ban.docx" ^
  --co-quan "TEN CO QUAN BAN HANH" ^
  --so-ky-hieu "12/CV-ABC" ^
  --dia-danh-ngay "Ha Noi, ngay 15 thang 4 nam 2026" ^
  --trich-yeu "V/v …" ^
  --body-pt 13 ^
  --line-exact-pt 16
```

- Long body: UTF-8 text file + `--noi-dung-file`.
- `--line-exact-pt 0` → single line spacing.
- `--body-pt 14` →14 pt body (default 13).

## Branch A — Technical specs (Decree 30 / common administrative)

| Item | Value |
|------|--------|
| Paper | A4 portrait (297 mm × 210 mm) |
| Font | Times New Roman, Unicode, **black** |
| Body size | **13 pt or 14 pt** (default **13** if unspecified) |
| Margins | Top **20 mm**, bottom **20 mm**, left **30 mm**, right **15 mm** |
| Line spacing | Single **or** Exactly **15–18 pt** (script default **16 pt**) |
| Paragraph spacing | **6 pt** before/after typical body paragraphs |
| Page numbers | **Centered** in **header**, Times New Roman, **13–14 pt** |

## Branch A — Front block (centered stack / Decree 30 style)

**If `Mau_cong_van_ND30_tai_ve.docx` exists:** ignore this “centered stack” list for **new** documents — the template’s **tables** define the layout. This list applies only to **legacy** output from `office_skill_cli.py legacy` or documents without the template.

If the document includes **Quoc huy** beside **Quoc hieu**, prefer the **Table layout for header** section above instead of a single centered column.

1. **State name** — full official name of the country, uppercase — **13 pt**, bold, **center**.
2. **Motto** — three standard hyphenated phrases — **13 pt**, bold, **center** (capitalization per Decree 30 practice).
3. **Issuing body** — **13 pt**, uppercase, bold, **center**.
4. **Number / symbol** — **13 pt**, regular, **center**.
5. **Place and date** — **13 pt**, italic, **right**.
6. **Subject line** — **13 pt**, regular, **center**.

Body after the subject: follow the technical table above unless the user specifies otherwise.

## Branch B — Legal normative documents (Decree 78/2025/NĐ-CP annex)

Summarized from the **Decree 78/2025/NĐ-CP** annex on form and presentation of legal normative instruments. For authoritative wording, use the official published annex and [reference-van-ban-quy-pham-phap-luat.md](reference-van-ban-quy-pham-phap-luat.md).

### I. General

| Item | Rule |
|------|------|
| Paper | A4 **210 mm × 297 mm** |
| Orientation | Along the **long** side of A4 (portrait); **appendices** with wide tables may use **landscape** |
| Margins | Top / bottom / right: **15–20 mm**; left: **30–35 mm** |
| Font | Times New Roman, Unicode **TCVN 6909:2001**, **black** |
| Sizes / styles | Per **Part III, Section 1** of the annex |
| Positions | Per **Part IV, Section 1** of the annex (diagram boxes **1–11**) |
| Page numbers | **Arabic numerals**, **13 pt**, upright; **page 1 not shown**; **horizontally centered** in the **top margin** area |
| Attached instruments | **Separate** page numbering **per** attached instrument |
| Appendices | **Separate** page numbering **per** appendix |

### II. Main front-matter (first page layout — **not** like branch A)

- **State name** (official full name of the country): **uppercase**, **12–13 pt**, upright, **bold**; **top of page, RIGHT** on page 1.
- **Motto** (three hyphenated phrases): **13–14 pt**, normal type with **capitalized** phrase starts, upright, **bold**; **centered** immediately **below** the state name; hyphens `-` between phrases, proper spacing; **full-width horizontal rule** below (solid line, length = line length).
- **Issuing body / authority**: **one line**, **uppercase**, **12–13 pt**, upright, **bold**; **top of page, LEFT** on page 1; **horizontal rule** below, length **1/3 to 1/2** of the text line, **centered** under the line. **Orders / President’s decisions**: **national emblem** above the title line.
- **Number and symbol**: **centered** under the issuing block; characters **contiguous** (no extra spaces); label **So** + colon rules, **13 pt** regular; symbol part **uppercase 13 pt**; slash `/` and hyphen `-` group rules; leading **0** for values &lt; 10; **Roman numerals** for National Assembly session key where required — follow annex in full.
- **Instrument title**: **Laws / ordinances** — **uppercase 14 pt**, bold, centered; type and title on **separate lines**. **Other instruments** — **type**: uppercase 14 pt bold centered; **short title**: regular **13–14 pt** bold centered below; **horizontal rule** under short title (1/3–1/2 line, centered). **Attached instruments**: parenthetical note **13–14 pt italic** centered under title; another rule below per annex.
- **Place + adoption/signature date** (except instruments of National Assembly, NA Standing Committee, People’s Council per annex): **same line as number/symbol**; **13–14 pt**, **italic**, comma after place; positioned **below**, **centered**, visually balanced with state name + motto.
- **Closing / signatures**: job titles **uppercase 13 pt** bold; **TM. / KT. / Q.** rules; signer name **13–14 pt** bold centered under title block; **left/right/balanced** layouts for joint resolutions as in annex.
- **Distribution block** (Noi nhan): label line **12 pt** italic bold, flush left aligned with signature block line; list **11 pt** regular upright; each addressee **on its own line**, leading `-`, trailing `;`; final **Luu:** line per annex.

### III. Body

- **Legal basis** (can cu ban hanh): **14 pt** regular **italic** under the title; **one basis per line**, line ends with **`;`**. Special centered headings (e.g. **QUYET NGH**, **QUYET DINH**, **LENH**, per annex) **uppercase 13 pt** bold, own line, **`:`**, centered — per instrument type in annex.
- **Main text**: **13–14 pt** regular upright, **justified**; **first line indent 1 cm to 1.27 cm** after line break; **≥6 pt** spacing between paragraphs; line spacing from **single** up to **1.5 lines**.
- **Structure**: **Part/Chapter** labels + **Roman** numbers, centered **13–14 pt** bold; **Part/Chapter names** below, **uppercase 14 pt** bold centered. **Section/Subsection**: **Arabic** numbering; names **uppercase 13–14 pt** bold centered. **Article (Dieu)**: **13–14 pt** bold, **left indent 1–1.27 cm**, **Arabic** article number + **`.`**. **Clause (khoan)**: **Arabic** + **`.`**. **Subpoints**: Vietnamese letters + `)`.
- **Appendices**: from **two** appendices upward, appendix numbers use **Roman numerals**; **Phu luc** line **13–14 pt** bold centered; appendix name **uppercase** bold; reference line under name **13–14 pt italic** centered.

### IV. Diagram (first A4 page)

Boxes **1–11** map to: motto block; issuing body; number/symbol; place & date; instrument title; body text; signature lines; seal; distribution; classification marking; typist mark / copy count — match the **official diagram** in Decree **78/2025/NĐ-CP**.

## Agent workflow

### Mandatory execution policy (always apply)

1. **Always export a `.docx` immediately** after gathering enough required inputs from the user/context.
2. **Do not wait for extra confirmation** like “co xuat file khong?” before saving output.
3. If output path is not explicitly provided, auto-create a deterministic file name in the same folder:
   - Edit existing file: `<ten_goc>_sua.docx` (unless user requests overwrite).
   - New file: `<ten_van_ban>.docx`.
4. After save, **only** run safe, template-preserving checks by default:
   - Header/signature layout tables keep correct widths and alignments (no table flattening).
   - Bottom-left `Noi nhan` block matches required font/weight/style/alignment rules **without rebuilding the block**.
5. For reply-letter body content (outside layout tables), enforce by default:
   - **Justified** paragraphs.
   - **14 pt** body font unless user requests 13 pt.
   - **Continuous flow**: remove empty spacer paragraphs (no blank line between consecutive body paragraphs).
   - Paragraph spacing: **Before = 0 pt, After = 0 pt**.
   - Special indentation: **First line = 1.27 cm**.
   - Before `Noi nhan` + signature block: keep **exactly 1 blank line** after the last body paragraph.
6. **Never keep raw placeholders** in final output:
   - Do not leave bracket markers like `[Tên đơn vị]`, `[Địa danh]`, `[Họ và tên]`.
   - Auto-fill placeholders with Vietnamese defaults (with diacritics). If no exact mapping exists, remove brackets and keep readable text.
7. Post-export gate (must pass):
   - `document.tables` count is **>= 2** for ND30 branch A outputs.
   - If table count is `< 2`, mark output invalid and regenerate from `Mau_cong_van_ND30_tai_ve.docx`.
8. Vietnamese text quality:
   - Preserve Vietnamese diacritics; do not convert content to non-diacritic ASCII.
   - For ND30 outputs, prefer Vietnamese default values with full accents when auto-filling missing fields.
   - If source body appears non-diacritic, **stop export and report error** instead of writing a broken output file.

### ND30 output quality guarantee (use the frozen template)

When the user wants a Decree-30-style reply letter and there is a “draft” `.docx` whose output looks bad:

- **Do not attempt to "fix" the draft into shape.** Rebuild the final output by **copying** the frozen template `Mau_cong_van_ND30_tai_ve.docx` and inserting the body text into it.
- Use helper script `scripts/office_skill_cli.py` subcommand `rebuild`:
  - It copies the template to the output file.
  - It removes any template body paragraphs between the **two layout tables**.
  - It inserts the source document’s body paragraphs **before** the signature table.
  - It removes empty spacer lines from source body, keeps paragraphs continuous, applies justify by default, and sets body font size default 14.
  - It fills common ND30 placeholders (header/signature/body) with Vietnamese defaults if input data is missing.
  - It sanitizes remaining placeholders globally so `[]` markers never appear in exported `.docx`.
  - It never changes the template tables’ structure.
  - It restores the **Straight Connector** horizontal line under the motto if that OOXML block is missing (same mechanism as `oxml_underline_after_motto.xml`).

**1. User edits an existing `.docx`**  
Clarify scope → edit in place → save; do not regenerate from the Decree 30 script.

**2. User wants a new branch A file**  
**Copy** `Mau_cong_van_ND30_tai_ve.docx` → Save As → fill **tables[0]**, **tables[1]**, and body text. **Do not** emit a header as plain paragraphs only.  
If output has **0 tables** or header is plain stacked paragraphs, treat it as failed generation and regenerate via template workflow.

**3. User wants a new branch B file**  
Do **not** use the Decree 30 script for the cover/first page. Build in Word or python-docx per **Branch B**; implement **different first page** for headers if needed so **page 1 has no page number**; page numbers **13 pt**, **top margin**, **centered** from page 2 onward (or as in the user’s template).

## Implementation notes (python-docx)

- Branch A: `Mm(...)` margins; `PAGE` field in header (centered).
- Branch B: often needs **“different first page”** header/footer behavior; align section breaks and fields with the annex; cross-check **margin bands** (15–20 mm top/bottom/right, 30–35 mm left).

## Personalization

Edit this `SKILL.md` or the reference file. Keep `description` rich in **keywords** (Decree 30, Decree 78, legal normative, luat, nghi dinh, thong tu, Word, docx, etc.) so the agent loads the skill when relevant.
