# M言語 (Power Query) — coding rules

Excel Power Query (M) tools for pharmacy operations. One subfolder per project; also read the project's CLAUDE.md.

## Runtime constraints
- Windows Excel (M365) Power Query only. **No VBA/macros, no external servers/DBs/add-ins.**
- Deployment = user opens a .txt and pastes each query into the Advanced Editor. Deliverable is one paste-ready `.txt` (UTF-8, CRLF) per project, built from `.m` sources (a project may add a second `.txt` for one-time setup queries, e.g. hand-entered tables, from its own manifest).
- Inputs are exports from the pharmacy system (エリシアS) as `.xls/.xlsx/.csv/.txt`. Every file is opened through `_shared/22_fx_ファイルを開く` (㉒), which decides the format by content (magic bytes), not by extension, auto-detects text encoding and delimiter, and raises a named error (`[読めないファイル] 「<file>」…`) for anything it cannot open.
- Header row is located by searching for a known label (e.g. `薬品名`, `品名`), not by position; in a workbook, pick the sheet by that label too (not "first sheet").
- Observed `.xls` behaviour in this Excel's Power Query (2026-11-07, replaces earlier guesses):
  - `Excel.Workbook` nav table has only `Name` and `Data` (no `Item`, `Kind`, `Hidden`); rows are sorted by name, not workbook order; hidden sheets are missing. Sheet names come without quotes or `$` (`4月10日(金)`), but defined names appear as rows like `'10月1日(水)$'Print_Area`, `'Sheet 2$'FilterDatabase`; a workbook-level name is a plain name indistinguishable from a sheet.
  - Every data column is text. Numbers arrive as displayed (`1,234,568`, `2.50`, `12.5%`, `¥9,801`; a `0` format rounds; a long number in General format arrives as `4.9996245439e+12` and cannot be recovered — JAN codes at risk), dates as displayed (`2026/10/01`, `2026/10/01 09:00:01`), booleans `TRUE`/`FALSE`, errors `#N/A`. Leading spaces are preserved; no cells were lost in mixed-type columns.
  - The third argument of `Excel.Workbook` makes no difference for .xls. `useHeaders = true` switches to typed columns and loses mixed-type cells — never use it.
  - Excel 5.0/95 files read like Excel 97-2003. CSV / tab text / HTML / XML Spreadsheet 2003 renamed to `.xls` fail in `Excel.Workbook` with `DataFormat.Error: 外部テーブルのフォーマットが正しくありません。` (㉒ reads text content itself and raises a named error for HTML/XML).
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
- `.m` files: UTF-8, LF. Built `.txt`: CRLF. Build with the project's `build.py`; it runs `_tools/mcheck.py` (bracket balance, top-level let bindings and identifier characters — not semantics).

## Shared queries (`_shared/`)
- Queries used by several projects live in `_shared/` and are referenced from each `manifest.txt`.
- Changing a shared query ⇒ rebuild every project using it.
- Current shared set: 00_fx_正規化, 01_設定値, 02_fx_使用量読込, 03_日次データ, 04_fx_卸の休み, 22_fx_ファイルを開く (㉒, created right after ⓪).
- The ordering tool PowerQuery_ABC (v0.2.4, not yet in this repo; there fx_卸の休み is ㉒) contains copies of the shared queries. Until it is migrated here, mirror any shared change there manually. Pending since 2026-11-07: the changed ①② and the new fx_ファイルを開く (a new query there) are not mirrored yet.

## Coding rules
- **Settings table holds only calculation variables (days, ratios, thresholds) and folder paths.** Fixed labels (export column names, file-name markers) are hard-coded in `01_設定値`. Read settings as `Num("項目", default)` — blank ⇒ default.
- Optional hand-entered tables: `try Table.SelectColumns(Excel.CurrentWorkbook(){[Name="X"]}[Content], cols, MissingField.UseNull) otherwise #table(cols, {})`. Tables are read by table name, not sheet name.
- Column-name matching: normalize with `fx_正規化`, strip spaces/newlines, unify parentheses to full-width.
- Drug names pass through `fx_正規化` (full-width ASCII→half, half-width kana→full). Keep 略称 raw (display only).
- Don't use reserved words as identifiers. Japanese identifiers are fine. In nested `each`, name the outer row with `(r) =>`.
- Identifiers (including record field names) may only use letters/digits/underscore (Japanese letters included); any other character such as `・` `（` `）` `／` must be quoted as `#"…"` — Excel parses the query as invalid otherwise (saved as `#!"…"`, evaluates to 「名前 'Source' を認識できません」). `_tools/mcheck.py` flags this as IDENT?.
- Open input files only through ㉒ `fx_ファイルを開く`; never call `Excel.Workbook`/`Csv.Document` on an export directly in another query.
- `.xls` and text files deliver numbers and dates as text, so every number/date conversion must be tested with text input (`1,234`, `100錠`, `¥9,801`, `▲50`, `12.5%`, `2026/10/01 09:00:00`), not only with typed `.xlsx` cells.
- Never put a cleaning expression that can raise inside `try … otherwise null` without testing the text path: the error is swallowed and every value silently becomes null (0.4.0 bug: `Text.Remove(txt, {… "mL", "cap"})` — `Text.Remove` accepts single-character strings only).
- M facts: `Text.FromBinary(sjisBytes, 65001)` does not raise, it yields U+FFFD; `Csv.Document` without `Columns` takes the column count from the first row (later wider rows are truncated), and with `Columns = n` fills missing cells with empty strings, not null.
- Use `Table.Buffer`/`List.Buffer` for reused lookups; avoid per-row `Table.SelectRows` over large tables.
- `Table.FromRecords` types: use `nullable`; apply `Int64.Type`/`Percentage.Type` afterwards via `Table.TransformColumnTypes`.
- Only raise errors for missing required settings and for input files that cannot be read (named error `読めないファイル` carrying the file name — never skip such a file silently), with an actionable Japanese message.
- Comments and all user-facing text in plain Japanese that pharmacy staff can read (avoid jargon/English abbreviations).
- Built `.txt` header must cover: version, changes, setup, how to create queries, load targets, operation, caveats.

## Versioning
- `major.minor.patch`: new query / table-column change = minor; fix = patch.
- On bump: update `.txt` header change list and the project's 引継ぎ手順書 history.
- Author name on documents: 関根禎浩.

## Verification
1. `python build.py` → `problems: 0` (mcheck checks bracket balance, top-level let bindings and identifier characters — not semantics).
2. Build the synthetic data in every input format and run each in Excel: `python tests/make_testdata.py` (`.xlsx`), `python tests/make_formats.py` (`.csv`/`.txt`), `pwsh -File tools\make_xls.ps1` (genuine `.xls` via Excel; dev PC only). Paste into Excel, Refresh All, compare with `tests/out/期待値.xlsx` for each format (a project may automate this, e.g. PowerPivot_ABC_Analyze `tools/` + `tests/compare_excel.py`; dev PC only). `.xlsx` alone is not verification — the store's files are `.xls`.
3. Passing the syntax check ≠ working. Say "not yet verified in Excel" until it is, and when it is, name the formats verified (e.g. "verified in Excel with .xlsx/.xls/.csv/.txt synthetic data"); real store exports count separately.
4. Replies and docs for the user are in Japanese.
