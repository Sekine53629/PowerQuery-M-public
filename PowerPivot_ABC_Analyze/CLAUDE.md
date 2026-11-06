# PowerPivot_ABC_Analyze — inventory analysis

Read `../CLAUDE.md` first. Project-specific rules only.

## Goal
Psychiatric-clinic-adjacent pharmacy (tablets only). Keep the stock that avoids shortages while minimizing tied-up capital and waste. Outputs support dead-stock handling, returns, transfers to other stores, and near-expiry monitoring. Daily ordering is out of scope (separate tool PowerQuery_ABC).

## Layers (keep separate)
1. **Data**: facts only (T_日次, T_月次, M_日付, M_薬品 attributes + metrics). No judgments.
2. **EDA**: choosing metrics and thresholds. Done in Excel pivots / Python; conclusions go to `docs/`. Not a query.
3. **Judgment**: thresholds live in the settings table; results are emitted as reason codes with messages (T_判定理由). Always keep the reason.

## Data semantics (do not confuse)
- **薬品使用量分析データ (実績 folder) = PLAN.** Built from next-visit schedules: scheduled (定期) patients only; excludes ad-hoc/new/other-dept/skipped visits. Treat past dates as plan too. 患者数 = scheduled frequency only.
- **出納 monthly = ACTUAL.** Lists only drugs with any movement that month. No movement ⇒ stock unchanged, so current stock = 在庫数量 of the last month the drug appeared.
- 予定とのずれ = 払出 − 予定; can be + (ad-hoc, new, re-visits, dose up) or − (discontinued, leftover adjustment, no-show). Monthly sums cancel within a month (treat magnitude as a lower bound).
- Drugs unmoved since before the oldest loaded 出納 month appear nowhere (needs a full stock list once).
- No YJ code in the plan/ledger exports (the 不動品リスト does carry YJ/JAN codes and 薬価). Key = normalized drug name (出納: 品名). 略称 (column B) is display-only.
- 品名 in 出納 and in the three system lists starts with a mark: `劇 ` (劇薬), `向 ` (向精神薬), etc.; drugs without a mark start with 3 spaces. Queries strip the mark and leading spaces before matching (⑤ keeps it as 品名の記号). The plan file's 薬品名 has no mark.
- File names in this section show the `.xls` form the store exports today; since 0.4.0 every input may be `.xls/.xlsx/.csv/.txt` (see "Input formats").
- Three optional system lists (structure reports 2026-11-03), header row 1, data from row 2, no merged cells, newest file in the folder is read:
  - `経過措置薬品リスト.xls` (経過措置フォルダパス): per drug. A 記号 (× = stock 0), B 製剤コード, C 製剤, D コード, E 品名, F メーカーコード, G メーカー, H 経過措置 (date), I 在庫数量, J 単位.
  - `YYYYMMDD_期限切れリスト.xls` (期限切迫フォルダパス): per purchase-slip line (lot). 26 columns incl. E 品名, H 棚番, I 伝票日付, N 仕入先, P 仕入数量, Q 単価, R 期限, S 期限切れ数量, T 単位, V ロット番号.
  - `YYYYMM_不動品リスト.xls` (不動品フォルダパス): per purchase-slip line (lot). 42 columns incl. E 品名, H 棚番, R 期限日, S 最終払出日, Z 数量, AA 単位, AC 薬価, AH ロット番号, AJ 基準日, AK 最終使用日, AL 不動日数, AM YJコード, AO JANコード.
- 出納 file: `yyyymmdd_yyyymmdd_数量一覧.xls`, one sheet, 10 fixed columns. Header row = the row containing 品名; rows above it may carry 出力日/期間. Period = first in-file date pair (前 ≦ 後) above the header, else the two 8-digit numbers in the file name. Real-file layout above the header is not yet confirmed with a sample. Analysis-data file (real name `薬品使用量予測データYYYYMM.xls`): prefix `薬品使用量`, daily sheets `10月1日(水)`, header row 1 merged with blank row 2, data from row 3. Confirmed layout (screenshot 2026-10-28): A 製剤 (`内服`), B 略称 (3 half-width kana + 2 digits, e.g. `ｱｺﾌ01`), C 薬品名, D 患者数, E 予定使用量, F 在庫数量 (incl. pending orders, = E + G), G 予定在庫数, H 単位 (`錠`/`g`/`cap`), hidden I 予定使用量（当日までの小計）, hidden J 在庫数量（発注残を含まない） (structure report 2026-10-09; `01_設定値` labels match: 在庫量 = J, 在庫量込み = F). 製剤 is the fixed string `内服` + 2 trailing spaces; `mg` in 薬品名 is full-width `ｍｇ` (fx_正規化 handles both). Sheets exist for every calendar day (store has no closing day). `期間すべて` is the plain column sum of the daily sheets (stock columns are summed too, so they are meaningless there). Confirmed layout (screenshot of 薬品使用量予測データ202510.xls, 2026-10-28): A 製剤 (`内服`), B 略称 (3 half-width kana + 2 digits, e.g. `ｱｺﾌ01`), C 薬品名, D 患者数, E 予定使用量, F 在庫数量, G 予定在庫数 (= F − E), H 単位 (`錠`/`g`/`cap`); columns I–J exist but are hidden (headers unconfirmed; `01_設定値` assumes I = `在庫数量（発注残を含まない）`). Data from row 3. Sheets exist for every calendar day (store has no closing day). `期間すべて` is the plain column sum of the daily sheets (stock columns are summed too, so they are meaningless there).

## Reference period and units (0.3.0)
- The settings table has no 開始日/終了日. `_shared/01_設定値` derives them: 終了日 = 予測開始 − 1 day (= 発注日 by default), 開始日 = first day of the month containing (終了日 − 1 year + 1 day). If a settings row 開始日/終了日 exists and is non-blank it wins (PowerQuery_ABC compatibility). ③ plan data and ⑤ 出納 both read only that window (⑤ keeps files whose period overlaps it).
- No pack size anywhere: no 包装単位 column, no 標準の包装単位, no 1箱消化日数/返品できる箱数. The only small-demand threshold is 極小需要の月の払出 (default 100), shared by ⑲ E, ⑬ G02 and ⑮ 持ち方 (回転 = 直近の1日平均払出 × 30 ≧ threshold). 発注量の目安 is plain days × per-day demand.

## Input formats (0.4.0)
- Accepted: `.xls/.xlsx/.csv/.txt(.tsv)`, may be mixed in one folder. ②⑤⑭⑳㉑㉓ open every file through ㉒ `fx_ファイルを開く(content, fileName, optional headerLabel)` (`_shared/22_fx_ファイルを開く.m`), which decides by content, not extension: bytes D0 CF 11 E0 → .xls, 50 4B → .xlsx, otherwise text. Returns シート名／Data／形式.
- Text files: encoding from BOM (UTF-8/UTF-16), else decode as UTF-8 and use 932 (Shift_JIS) if U+FFFD appears; delimiter = tab or comma, whichever is more frequent in the first 50 non-empty lines; column count from the longest line; all values text, blanks null.
- `headerLabel` (`"品名"` for 出納 and the three lists) returns only the sheets whose first 30 rows contain that header — `.xls` sheets arrive sorted by name, so "first sheet" means nothing.
- Errors: an empty file, content starting with `<` (HTML/XML saved as .xls) or any open failure raises a named error (`[読めないファイル] 「<file>」を開けません：…`). ⑤⑭⑳㉑ raise the same kind of error when a file in the reading window has no 品名 header row (no silent skip). ② still skips a sheet without a 薬品名 header.
- Plan data (②) in text files has no sheets, so one of two shapes is read, otherwise the file is not read: (1) a header column `日付` (label fixed in ①), any number of days per file; (2) one file per day with an 8-digit date in the file name (`薬品使用量分析データ20261010.csv`). **Provisional: the real エリシアS CSV layout of this report is not confirmed.**
- 出納 and the lists in text files: same header names as the .xls (header row found by 品名); 出納's in-file period lines above the header still work.
- Number conversion (②⑤⑭⑳㉑): remove spaces, thousands separators, yen marks; ▲/△ → minus; `Number.FromText`, else the leading numeric part (`1,234`, `100錠`, `¥9,801`, `▲50`, `12.5%`). Date reading (②⑭⑳㉑): date cells, `yyyy/mm/dd`, `yyyy-mm-dd`, `yyyy年m月d日`, 8-digit `yyyymmdd`, text with a time part.
- ㉓ `読込の確認` (`src/23_読込の確認.m`, loaded to sheet `読込の確認`): row 1 = reading period (開始日～終了日, 発注日); then one row per file in each of the five folders (実績／出納／期限切れリスト／経過措置薬品リスト／不動品リスト) with 種類, ファイル名, 更新日時, 形式, シート数, 見出しの行, 行数, 結果. 結果 = `読める…` / `読まない：…` (not a target: wrong name or extension, outside the period, a newer list exists, no date for a text plan file) / `読めない：…` (cannot open, no header), each with the reason. When numbers are missing, look here first; at the store, this sheet's content is what gets reported.
- `bi/model.py` (Python reference) follows the same rules: `open_tables()` / `first_table()` / `list_files()`, `.xls` via xlrd.

## Tables (data model)
| Table | Grain | Keys |
|---|---|---|
| T_日次 | date × drug | 日付, 薬品名 |
| T_月次 | month × drug | 月, 薬品名 |
| M_日付 | date | 日付 |
| M_薬品 | drug | 薬品名 |
| M_理由コード | code | コード |
| T_判定理由 | drug × code | 薬品名, コード |

Relationships: T_日次/T_月次 → M_日付, M_薬品; T_判定理由 → M_薬品, M_理由コード.
- Compute everything in Power Query; **no DAX measures**. Pivots should only need sums/counts.
- When adding a metric, document it in the header comment of `09_M_薬品.m`.

## Build
- Current version 0.4.0: 24 queries ⓪–㉓. `build.py` with no arguments builds both manifests: `manifest.txt` (main queries; ㉒ right after ⓪, ㉓ last → `dist/PowerPivot_ABC_<ver>.txt`) and `manifest_init.txt` (hand-entered tables, unchanged in 0.4.0 → `dist/PowerPivot_初期テーブル作成_<ver>.txt`). Keep both VERSION lines equal.
- Sources: ⓪–④ and ㉒ in `../_shared/`, ⑤–㉑ and ㉓ in `src/`. Upgrading a 0.3.0 workbook: re-paste ①②⑤⑭⑳㉑, create ㉒ and ㉓.
- Hand-entered table definitions (設定, 経過措置リスト, 発注除外, 在庫情報, 追加休業日, 薬品グループ) live in `src/init/` as `<A–F>_<table name>.m` with header mark `A`…`F` instead of a circled number, because the user deletes these queries after unlinking the tables; they are not part of the ⓪–㉓ sequence. Adding a column to one of them is a minor bump and must also be reflected in the table list of `src/_header.txt`.
- `dist/*.txt` is committed (it is the deliverable); only the current version is kept there.
- Shared queries ①② and ㉒ changed in 0.4.0: mirror them to PowerQuery_ABC manually (㉒ is a new query there). Not done/verified yet.
- Excel-side verification is automated by `tools/excel_install.ps1` (install queries, connections, relationships, the `読込の確認` sheet, refresh; `-SettingsOnly` rewrites only the settings folders), `tools/excel_dump.ps1` (dump query results to CSV; `-Touch`; `-Queries a,b,c` also works with `pwsh -File`) + `tests/compare_excel.py` (compare with `tests/out/期待値.xlsx`). Dev PC only; never on the store PC. The prototype workbook `prototype/ABC分析.xlsx` is untracked.
- Verification must cover all four formats, never `.xlsx` alone (steps in `引継ぎ手順書.md` section 3): `python tests/make_testdata.py` (`.xlsx` under `tests/out`; numbers ≥ 1000 get a `#,##0` display format so `.xls` delivers `1,700` as text; rebuilding deletes `tests/out/形式別`, so re-run the next two) → `python tests/make_formats.py` (`形式別/csv`: comma, Shift_JIS, plan data one file per day; `形式別/txt`: tab, plan data monthly with a `日付` column, UTF-8 with BOM / ledger UTF-8 without BOM / lists UTF-16) → `pwsh -File tools\make_xls.ps1` (genuine `.xls` under `形式別/xls` via Excel, no LibreOffice; dev PC only) → `python tests/make_formats.py --check` (Python reference: xls/csv/txt vs xlsx) → install → per format: `excel_install.ps1 -SettingsOnly` with that format's folders, dump, compare.

## Excel verification facts
2026-11-07 (0.4.0, dev PC, synthetic `tests/out`: 6 months, 14 drugs):
- Why 0.4.0: on real store data not a single file was read, and up to 0.3.0 Excel verification had used `.xlsx` only — `.xls` had never been run. Reproduced with genuine `.xls`: in `.xls` every cell arrives as text, and ②'s number converter called `Text.Remove(txt, {… "mL", "cap"})`; `Text.Remove` takes single characters only, the error (`値が 1 文字の文字列ではありません。`) was swallowed lazily by `try Number.FromText(cleaned) otherwise null`, so every text number became null → 日次データ 0 rows (647 expected), M_薬品 13 (14), T_月次 58 (67). `.xlsx` cells are numbers and took the other branch. In the shared query since the first commit. 出納 and the three lists read fine from synthetic `.xls`. **Whether this bug alone explains the store symptom is not confirmed.**
- Result: 0.4.0 installed with no invalid-formula error; for each of .xlsx, .xls, .csv, .txt: 日次データ 647, 出納 58, 経過措置該当 2, 不動品リスト 5, T_ロット 8, M_薬品 14, T_月次 67, T_判定理由 44; `tests/compare_excel.py` 違い 0; source-query dumps value-identical across formats; `読込の確認` showed the right 形式 (`Excel（.xls）`, `文字（カンマ区切り・Shift_JIS）`, `文字（タブ区切り・UTF-8）`, `文字（タブ区切り・UTF-16）`). Python reference identical for xls/csv/txt.
- Negative tests: an HTML file named `.xls` and a 0-byte file inside the reading period make 出納／日次データ fail with a message naming the file, while `読込の確認` still loads and shows the reason.
- The last small change (named error for a missing 品名 header, wording order in `読込の確認`): re-verified in Excel on all four formats (2026-11-07, differences 0); the missing-header error was triggered for 出納 and 不動品リスト only, not for the other two lists.
- **Not verified**: real store exports (.xls or .csv), `tests/out_clinic` (1 year, 200 drugs) vs `bi/out/モデル.xlsx`, PowerQuery_ABC with the changed shared queries.
- Timing on this small data: full install 110 s, すべて更新 82 s (0.3.0 measured ~5 min to load the 9 model tables and 3–4 min per すべて更新). Refresh is still slow on larger data: Power Query re-evaluates the dependency chain per loaded table, so ③ 日次データ (opens the files) runs 10+ times per refresh. Planned fix: stage ③ and ⑤ once (sheet or model) and read downstream from there.
- `.xls` behaviour observed in this Excel is listed in `../CLAUDE.md` ("Observed `.xls` behaviour"). Project consequences: JAN codes in the 不動品リスト are at risk (a long number in General format arrives as `4.9996245439e+12`, unrecoverable); leading spaces of 品名 are preserved.
- `$wb.Queries.FastCombine = $true` in a COM session lets queries run on workbook copies without the `-Touch` workaround.

2026-11-04 (0.3.0, 22 queries; still true):
- 0.3.0 matched `tests/out` in Excel — with `.xlsx` inputs only (superseded by the 0.4.0 run above).
- Loading a query into the data model from COM must use `Connections.Add2` with CmdType 6 (table collection) and command `"<query name>"`; the SQL form hangs Excel.
- Reopening a workbook whose queries were rewritten loses the "ignore privacy levels" grant → Formula.Firewall error on ③ 日次データ. Re-grant in Excel (クエリのオプション → 現在のブック → プライバシー → 無視) and save; rewriting any one query in the same session also clears it (`excel_dump.ps1 -Touch`). Some cases are not fixed by re-granting; cause unknown (check the global privacy setting).
- Identifiers may only contain letters/digits/underscore; `・`, `（）`, `／` etc. must be quoted as `#"…"` (⑱ `切替・転送の案` broke in Excel: saved as `#!"…"`, evaluates to 「名前 'Source' を認識できません」). `_tools/mcheck.py` flags this as IDENT?.

## Reason codes
- Codes are permanent (never renumber/reuse; retire as 欠番 — keep the row in ⑩ with 種類 "欠番", never emit it in ⑪). Retired: S03 (0.3.0, pack-based). Prefixes: S stock/dead, D demand, P plan gap, E expiry/regulatory, Q data quality, J final verdict (not implemented).
- Messages: polite Japanese with numbers via `{値}{線}{在庫}` placeholders, e.g. 「{値}ヶ月、不動です。（在庫 {在庫}）」.
- Thresholds are settings items read with `Num("項目", provisional default)` in `11_T_判定理由.m`. Defaults are provisional until EDA sets them.

## Status / next steps
See `引継ぎ手順書.md` (Japanese).
