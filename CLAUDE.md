# M言語 (Power Query) — coding rules

Excel Power Query (M) tools for pharmacy operations. One subfolder per project; also read the project's CLAUDE.md.

## Runtime constraints
- Windows Excel (M365) Power Query only. **No VBA/macros, no external servers/DBs/add-ins.**
- Deployment = user opens a .txt and pastes each query into the Advanced Editor. Deliverable is one paste-ready `.txt` (UTF-8, CRLF) per project, built from `.m` sources (a project may add a second `.txt` for one-time setup queries, e.g. hand-entered tables, from its own manifest).
- Inputs are .xls/.xlsx exports from the pharmacy system (エリシアS). Legacy .xls quirks to always handle:
  - `Excel.Workbook` nav table may lack the `Kind` column
  - sheet names may carry `'` and a trailing `$`; defined names (`Print_Area`, `_xlnm`, names with inner `$`) appear as rows
  - header row is located by searching for a known label (e.g. `薬品名`), not by position
- User sets privacy level to "ignore" for the workbook (Formula.Firewall).

## Data handling (hard rules)
- Never place, commit, or transmit real store data. Tests use only seeded synthetic data (`tests/`).
- No patient-level data; drug-level aggregates only.
- Do not propose or investigate direct DB access to the pharmacy system.

## Files
- One query per file: `NN_<query name>.m`. Query name must equal the Excel query name (referenced as `#"name"`).
- Each `.m` starts with this header (build/check scripts split on it):
  ```
  // ================================================================================
  // ⑤ クエリ名：<query name>（<load target>）
  //    <plain-Japanese description>
  // ================================================================================
  ```
- `.m` files: UTF-8, LF. Built `.txt`: CRLF. Build with the project's `build.py`; it runs `_tools/mcheck.py` (bracket balance + top-level let bindings only — not semantics).

## Shared queries (`_shared/`)
- Queries used by several projects live in `_shared/` and are referenced from each `manifest.txt`.
- Changing a shared query ⇒ rebuild every project using it.
- Current shared set: 00_fx_正規化, 01_設定値, 02_fx_使用量読込, 03_日次データ, 04_fx_卸の休み.
- The ordering tool PowerQuery_ABC (v0.2.4, not yet in this repo; there fx_卸の休み is ㉒) contains identical copies. Until it is migrated here, mirror any shared change there manually.

## Coding rules
- **Settings table holds only calculation variables (days, ratios, thresholds) and folder paths.** Fixed labels (export column names, file-name markers) are hard-coded in `01_設定値`. Read settings as `Num("項目", default)` — blank ⇒ default.
- Optional hand-entered tables: `try Table.SelectColumns(Excel.CurrentWorkbook(){[Name="X"]}[Content], cols, MissingField.UseNull) otherwise #table(cols, {})`. Tables are read by table name, not sheet name.
- Column-name matching: normalize with `fx_正規化`, strip spaces/newlines, unify parentheses to full-width.
- Drug names pass through `fx_正規化` (full-width ASCII→half, half-width kana→full). Keep 略称 raw (display only).
- Don't use reserved words as identifiers. Japanese identifiers are fine. In nested `each`, name the outer row with `(r) =>`.
- Identifiers (including record field names) may only use letters/digits/underscore (Japanese letters included); any other character such as `・` `（` `）` `／` must be quoted as `#"…"` — Excel parses the query as invalid otherwise (saved as `#!"…"`, evaluates to 「名前 'Source' を認識できません」). `_tools/mcheck.py` flags this as IDENT?.
- Use `Table.Buffer`/`List.Buffer` for reused lookups; avoid per-row `Table.SelectRows` over large tables.
- `Table.FromRecords` types: use `nullable`; apply `Int64.Type`/`Percentage.Type` afterwards via `Table.TransformColumnTypes`.
- Only raise errors for missing required settings, with an actionable Japanese message.
- Comments and all user-facing text in plain Japanese that pharmacy staff can read (avoid jargon/English abbreviations).
- Built `.txt` header must cover: version, changes, setup, how to create queries, load targets, operation, caveats.

## Versioning
- `major.minor.patch`: new query / table-column change = minor; fix = patch.
- On bump: update `.txt` header change list and the project's 引継ぎ手順書 history.
- Author name on documents: 関根禎浩.

## Verification
1. `python build.py` → `problems: 0` (mcheck checks bracket balance, top-level let bindings and identifier characters — not semantics).
2. `python tests/make_testdata.py --xls`, paste into Excel, Refresh All, compare with `tests/out/期待値.xlsx` (a project may automate this, e.g. PowerPivot_ABC_Analyze `tools/` + `tests/compare_excel.py`; dev PC only).
3. Passing the syntax check ≠ working. Say "not yet verified in Excel" until it is.
4. Replies and docs for the user are in Japanese.
