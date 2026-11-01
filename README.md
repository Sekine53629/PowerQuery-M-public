# PowerQuery-M

薬局業務向けの Excel Power Query（M 言語）ツール集。
エリシアS のエクスポート（.xls）を読み、在庫・発注の判断材料を作る。
VBA・マクロ・外部サーバーは使わず、貼り付け用の `.txt` を Excel の詳細エディターに貼るだけで動く。

## フォルダ

```
M言語/
  CLAUDE.md                 M言語の記述ルール（Claude Code 用・英語）
  _tools/mcheck.py          簡易構文チェック（括弧・let の束縛）
  _shared/                  複数プロジェクト共通のクエリ（00～04）
  PowerPivot_ABC_Analyze/   在庫分析（データモデル＋判定理由）… 引継ぎ手順書.md を参照
```

## セットアップ

- Python 3 と openpyxl（`pip install openpyxl`）
- Windows Excel（M365）

## 使い方

各プロジェクトのフォルダで `python build.py` を実行すると `dist/` に貼り付け用の `.txt` ができる。
詳細はプロジェクトの `引継ぎ手順書.md` を参照。

## データの取り扱い

- 実データ（店舗の出納・使用量）はリポジトリに置かない。テストは `tests/make_testdata.py` の架空データだけを使う。
- Claude に読ませるデータは各プロジェクトの `input/` に置く（中身は Git に記録されない）。
