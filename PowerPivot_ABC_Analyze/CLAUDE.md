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
- No YJ code in any export. Key = normalized drug name (出納: 品名). 略称 (column B) is display-only.
- 出納 file: `yyyymmdd_yyyymmdd_数量一覧.xls`, one sheet, header row 1, 10 fixed columns. Analysis-data file: prefix `薬品使用量`, daily sheets `10月1日(水)`, header row 1 merged with blank row 2.

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
- `build.py` with no arguments builds both manifests: `manifest.txt` (main queries ⓪–⑪ → `dist/PowerPivot_ABC_<ver>.txt`) and `manifest_init.txt` (hand-entered tables → `dist/PowerPivot_初期テーブル作成_<ver>.txt`). Keep both VERSION lines equal.
- Hand-entered table definitions (設定, 経過措置リスト, 発注除外, 在庫情報, 追加休業日, 薬品グループ) live in `src/init/` as `<A–F>_<table name>.m` with header mark `A`…`F` instead of a circled number, because the user deletes these queries after unlinking the tables; they are not part of the ⓪–⑪ sequence. Adding a column to one of them is a minor bump and must also be reflected in the table list of `src/_header.txt`.
- `dist/*.txt` is committed (it is the deliverable); only the current version is kept there.

## Reason codes
- Codes are permanent (never renumber/reuse; retire as 欠番). Prefixes: S stock/dead, D demand, P plan gap, E expiry/regulatory, Q data quality, J final verdict (not implemented).
- Messages: polite Japanese with numbers via `{値}{線}{在庫}` placeholders, e.g. 「{値}ヶ月、不動です。（在庫 {在庫}）」.
- Thresholds are settings items read with `Num("項目", provisional default)` in `11_T_判定理由.m`. Defaults are provisional until EDA sets them.

## Status / next steps
See `引継ぎ手順書.md` (Japanese).
