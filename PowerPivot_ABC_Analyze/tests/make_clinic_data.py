"""精神科門前薬局を想定した架空の処方データを作る（すべて架空。実データは使わない）
使い方（tests フォルダで）：
  python make_clinic_data.py --from 2026-09-01 --to 2026-09-14 [--per-day 175]
  出力：tests/out_clinic/<from>_<to>/
    処方_<from>_<to>.xlsx   … 処方明細・処方箋一覧・日別集計・薬品別集計・日別_予定と実績・薬品マスタ
    処方明細.csv            … 処方明細（UTF-8 BOM）
    日別_予定と実績.csv     … 日付×薬品の予定（次回予約から立つ量）と実績（実際に出た量）

想定：
  ・1日の処方箋 150～250 枚、ほとんどの患者は処方箋 1 枚（精神科）。内科など 2 枚目は少数
  ・定期の来局が大半（2週・4週・6週・8週・12週）。新患・臨時・再来が少し混じる
  ・睡眠薬など 30 日制限の薬があれば、処方日数は 28 日まで
  ・残薬調整（1 剤だけ日数を減らす）が少し入る
  ・薬は「成分規格」と「需要のまとまり」を持ち、同じ成分で複数の屋号（先発・後発）がある
"""
import os, sys, csv, math, random, argparse, datetime as dt
from collections import defaultdict, Counter
from openpyxl import Workbook
from openpyxl.styles import Font

SEED = 20261103
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out_clinic")
WD = "月火水木金土日"

# ------------------------------------------------------------------------------
# 薬品マスタ
#   cls : AD 抗うつ／AP 抗精神病／MS 気分安定／AX 抗不安／HY 睡眠／ADHD／DEM 認知症／AC 副作用止め／GI 内科系
#   gen : 先発／後発
#   pack: 包装単位（錠）　doses: 1日量の候補（錠）　w: 処方される重み（同じ成分の屋号シェアを含む）
#   limit: 投与日数の上限（30 日制限の向精神薬）
# ------------------------------------------------------------------------------
DRUGS = []


def D(name, cls, group, gen, pack, doses, w, unit="錠", limit=None, cluster=None):
    DRUGS.append(dict(name=name, cls=cls, group=group, cluster=cluster or group, gen=gen, pack=pack,
                      doses=[float(x) for x in str(doses).split(",")], w=w, unit=unit, limit=limit))


# ---- 抗うつ薬 ----
D("セルトラリン錠25mg「アメル」",          "AD", "セルトラリン25mg",     "後発", 100, "1,2",   28)
D("セルトラリン錠25mg「トーワ」",          "AD", "セルトラリン25mg",     "後発", 100, "1,2",   8)
D("ジェイゾロフト錠25mg",                 "AD", "セルトラリン25mg",     "先発", 100, "1,2",   2)
D("セルトラリン錠50mg「アメル」",          "AD", "セルトラリン50mg",     "後発", 100, "1,2",   22)
D("セルトラリン錠50mg「トーワ」",          "AD", "セルトラリン50mg",     "後発", 100, "1,2",   6)
D("エスシタロプラム錠10mg「DSEP」",        "AD", "エスシタロプラム10mg", "後発", 100, "1,2",   26)
D("レクサプロ錠10mg",                     "AD", "エスシタロプラム10mg", "先発", 100, "1,2",   12)
D("エスシタロプラム錠20mg「DSEP」",        "AD", "エスシタロプラム20mg", "後発", 100, "1",     10)
D("レクサプロ錠20mg",                     "AD", "エスシタロプラム20mg", "先発", 100, "1",     4)
D("パロキセチン錠10mg「日医工」",          "AD", "パロキセチン10mg",     "後発", 100, "1,2",   10)
D("パロキセチン錠20mg「日医工」",          "AD", "パロキセチン20mg",     "後発", 100, "1,2",   8)
D("パロキセチン錠20mg「トーワ」",          "AD", "パロキセチン20mg",     "後発", 100, "1,2",   3)
D("パキシル錠20mg",                       "AD", "パロキセチン20mg",     "先発", 100, "1,2",   2)
D("パキシルCR錠12.5mg",                   "AD", "パロキセチンCR12.5mg", "先発", 100, "1,2",   5)
D("フルボキサミンマレイン酸塩錠25mg「アメル」", "AD", "フルボキサミン25mg", "後発", 100, "2",   5)
D("フルボキサミンマレイン酸塩錠50mg「アメル」", "AD", "フルボキサミン50mg", "後発", 100, "2",   6)
D("デュロキセチンカプセル20mg「トーワ」",  "AD", "デュロキセチン20mg",   "後発", 100, "1,2",   14, unit="cap")
D("デュロキセチンカプセル30mg「トーワ」",  "AD", "デュロキセチン30mg",   "後発", 100, "1,2",   10, unit="cap")
D("ベンラファキシンSRカプセル37.5mg「サワイ」", "AD", "ベンラファキシンSR37.5mg", "後発", 100, "1,2", 6, unit="cap")
D("ベンラファキシンSRカプセル75mg「サワイ」",   "AD", "ベンラファキシンSR75mg",   "後発", 100, "1,2", 8, unit="cap")
D("イフェクサーSRカプセル75mg",           "AD", "ベンラファキシンSR75mg", "先発", 100, "1,2",  3, unit="cap")
D("ミルタザピン錠15mg「サワイ」",          "AD", "ミルタザピン15mg",     "後発", 100, "1,2",   18)
D("ミルタザピン錠30mg「サワイ」",          "AD", "ミルタザピン30mg",     "後発", 100, "1",     8)
D("リフレックス錠15mg",                   "AD", "ミルタザピン15mg",     "先発", 100, "1,2",   3)
D("トリンテリックス錠10mg",               "AD", "ボルチオキセチン10mg", "先発", 100, "1,2",   14)
D("トリンテリックス錠20mg",               "AD", "ボルチオキセチン20mg", "先発", 100, "1",     6)
D("トラゾドン塩酸塩錠25mg「アメル」",      "AD", "トラゾドン25mg",       "後発", 100, "1,2,4", 12)
D("トラゾドン塩酸塩錠50mg「アメル」",      "AD", "トラゾドン50mg",       "後発", 100, "1,2",   8)
D("アミトリプチリン塩酸塩錠10mg「サワイ」", "AD", "アミトリプチリン10mg", "後発", 100, "1,2,3", 5)
D("アミトリプチリン塩酸塩錠25mg「サワイ」", "AD", "アミトリプチリン25mg", "後発", 100, "1,2",   3)
D("アナフラニール錠25mg",                 "AD", "クロミプラミン25mg",   "先発", 100, "1,2,3", 2)
D("テトラミド錠10mg",                     "AD", "ミアンセリン10mg",     "先発", 100, "1,2,3", 2)
D("スルピリド錠50mg「アメル」",            "AD", "スルピリド50mg",       "後発", 100, "1,2,3", 10)
D("スルピリド錠100mg「アメル」",           "AD", "スルピリド100mg",      "後発", 100, "2,3",   3)
# ---- 抗精神病薬 ----
D("アリピプラゾール錠3mg「明治」",         "AP", "アリピプラゾール3mg",  "後発", 100, "1,2",   22)
D("アリピプラゾールOD錠3mg「トーワ」",     "AP", "アリピプラゾール3mg",  "後発", 100, "1,2",   6)
D("アリピプラゾール錠6mg「明治」",         "AP", "アリピプラゾール6mg",  "後発", 100, "1,2",   14)
D("アリピプラゾール錠12mg「明治」",        "AP", "アリピプラゾール12mg", "後発", 100, "1,2",   8)
D("エビリファイ錠3mg",                    "AP", "アリピプラゾール3mg",  "先発", 100, "1,2",   2)
D("レキサルティ錠1mg",                    "AP", "ブレクスピプラゾール1mg", "先発", 100, "1,2", 10)
D("レキサルティ錠2mg",                    "AP", "ブレクスピプラゾール2mg", "先発", 100, "1",   7)
D("レキサルティOD錠1mg",                  "AP", "ブレクスピプラゾール1mg", "先発", 100, "1,2", 3)
D("オランザピン錠2.5mg「DSEP」",           "AP", "オランザピン2.5mg",    "後発", 100, "1,2",   10)
D("オランザピン錠5mg「DSEP」",             "AP", "オランザピン5mg",      "後発", 100, "1,2",   12)
D("オランザピン錠10mg「DSEP」",            "AP", "オランザピン10mg",     "後発", 100, "1,2",   8)
D("オランザピンOD錠5mg「トーワ」",         "AP", "オランザピン5mg",      "後発", 100, "1,2",   4)
D("クエチアピン錠25mg「アメル」",          "AP", "クエチアピン25mg",     "後発", 100, "1,2,3", 16)
D("クエチアピン錠100mg「アメル」",         "AP", "クエチアピン100mg",    "後発", 100, "1,2,3", 10)
D("クエチアピン錠200mg「アメル」",         "AP", "クエチアピン200mg",    "後発", 100, "1,2,3", 5)
D("ビプレッソ徐放錠150mg",                "AP", "クエチアピン徐放150mg", "先発", 100, "1,2",  3)
D("リスペリドン錠1mg「アメル」",           "AP", "リスペリドン1mg",      "後発", 100, "1,2",   12)
D("リスペリドン錠2mg「アメル」",           "AP", "リスペリドン2mg",      "後発", 100, "1,2",   6)
D("リスペリドンOD錠1mg「アメル」",         "AP", "リスペリドン1mg",      "後発", 100, "1,2",   3)
D("パリペリドン徐放錠6mg「トーワ」",       "AP", "パリペリドン6mg",      "後発", 100, "1",     5)
D("インヴェガ錠6mg",                      "AP", "パリペリドン6mg",      "先発", 100, "1",     2)
D("インヴェガ錠9mg",                      "AP", "パリペリドン9mg",      "先発", 100, "1",     2)
D("ラツーダ錠20mg",                       "AP", "ルラシドン20mg",       "先発", 100, "1,2",   8)
D("ラツーダ錠40mg",                       "AP", "ルラシドン40mg",       "先発", 100, "1",     6)
D("シクレスト舌下錠5mg",                  "AP", "アセナピン5mg",        "先発", 100, "2",     2)
D("ロナセン錠4mg",                        "AP", "ブロナンセリン4mg",    "先発", 100, "2",     3)
D("ブロナンセリン錠4mg「DSEP」",           "AP", "ブロナンセリン4mg",    "後発", 100, "2",     4)
D("ブロナンセリン錠8mg「DSEP」",           "AP", "ブロナンセリン8mg",    "後発", 100, "2",     3)
D("ペロスピロン塩酸塩錠4mg「アメル」",     "AP", "ペロスピロン4mg",      "後発", 100, "2,3",   3)
D("ペロスピロン塩酸塩錠8mg「アメル」",     "AP", "ペロスピロン8mg",      "後発", 100, "2,3",   2)
D("ハロペリドール錠0.75mg「アメル」",      "AP", "ハロペリドール0.75mg", "後発", 100, "1,2",   3)
D("ハロペリドール錠1.5mg「アメル」",       "AP", "ハロペリドール1.5mg",  "後発", 100, "1,2",   3)
D("コントミン糖衣錠12.5mg",               "AP", "クロルプロマジン12.5mg", "先発", 100, "1,2,3", 3)
D("コントミン糖衣錠25mg",                 "AP", "クロルプロマジン25mg", "先発", 100, "1,2,3", 3)
D("ヒルナミン錠5mg",                      "AP", "レボメプロマジン5mg",  "先発", 100, "1,2",   4)
D("ヒルナミン錠25mg",                     "AP", "レボメプロマジン25mg", "先発", 100, "1,2",   2)
D("ロドピン錠25mg",                       "AP", "ゾテピン25mg",         "先発", 100, "1,2,3", 2)
D("クロザリル錠25mg",                     "AP", "クロザピン25mg",       "先発", 100, "2,4",   1)
D("クロザリル錠100mg",                    "AP", "クロザピン100mg",      "先発", 100, "2,3,4", 1)
# ---- 気分安定薬 ----
D("炭酸リチウム錠100mg「アメル」",         "MS", "炭酸リチウム100mg",    "後発", 100, "2,4",   6)
D("炭酸リチウム錠200mg「アメル」",         "MS", "炭酸リチウム200mg",    "後発", 100, "2,3,4", 10)
D("リーマス錠200mg",                      "MS", "炭酸リチウム200mg",    "先発", 100, "2,3,4", 3)
D("バルプロ酸ナトリウム徐放錠A100mg「トーワ」", "MS", "バルプロ酸徐放100mg", "後発", 100, "2,4", 4)
D("バルプロ酸ナトリウム徐放錠A200mg「トーワ」", "MS", "バルプロ酸徐放200mg", "後発", 100, "2,3,4", 8)
D("デパケンR錠200mg",                     "MS", "バルプロ酸徐放200mg",  "先発", 100, "2,3,4", 3)
D("ラモトリギン錠25mg「アメル」",          "MS", "ラモトリギン25mg",     "後発", 100, "1,2,4", 7)
D("ラモトリギン錠100mg「アメル」",         "MS", "ラモトリギン100mg",    "後発", 100, "1,2",   8)
D("ラミクタール錠25mg",                   "MS", "ラモトリギン25mg",     "先発", 100, "1,2,4", 2)
D("カルバマゼピン錠100mg「アメル」",       "MS", "カルバマゼピン100mg",  "後発", 100, "2,3",   3)
D("カルバマゼピン錠200mg「アメル」",       "MS", "カルバマゼピン200mg",  "後発", 100, "2,3",   4)
# ---- 抗不安薬 ----
D("エチゾラム錠0.5mg「トーワ」",           "AX", "エチゾラム0.5mg",      "後発", 100, "1,2,3", 20, limit=30)
D("エチゾラム錠1mg「トーワ」",             "AX", "エチゾラム1mg",        "後発", 100, "1,2,3", 10, limit=30)
D("デパス錠0.5mg",                        "AX", "エチゾラム0.5mg",      "先発", 100, "1,2,3", 4, limit=30)
D("ロラゼパム錠0.5mg「サワイ」",           "AX", "ロラゼパム0.5mg",      "後発", 100, "1,2,3", 16)
D("ロラゼパム錠1mg「サワイ」",             "AX", "ロラゼパム1mg",        "後発", 100, "1,2,3", 6)
D("ワイパックス錠0.5mg",                  "AX", "ロラゼパム0.5mg",      "先発", 100, "1,2,3", 3)
D("アルプラゾラム錠0.4mg「トーワ」",       "AX", "アルプラゾラム0.4mg",  "後発", 100, "1,2,3", 14)
D("アルプラゾラム錠0.8mg「トーワ」",       "AX", "アルプラゾラム0.8mg",  "後発", 100, "1,2,3", 4)
D("ソラナックス錠0.4mg",                  "AX", "アルプラゾラム0.4mg",  "先発", 100, "1,2,3", 3)
D("ジアゼパム錠2mg「トーワ」",             "AX", "ジアゼパム2mg",        "後発", 100, "1,2,3", 6)
D("ジアゼパム錠5mg「トーワ」",             "AX", "ジアゼパム5mg",        "後発", 100, "1,2,3", 4)
D("クロチアゼパム錠5mg「トーワ」",         "AX", "クロチアゼパム5mg",    "後発", 100, "1,2,3", 8)
D("クロチアゼパム錠10mg「トーワ」",        "AX", "クロチアゼパム10mg",   "後発", 100, "1,2,3", 4)
D("ブロマゼパム錠2mg「サンド」",           "AX", "ブロマゼパム2mg",      "後発", 100, "1,2,3", 3)
D("ブロマゼパム錠5mg「サンド」",           "AX", "ブロマゼパム5mg",      "後発", 100, "1,2,3", 2)
D("メイラックス錠1mg",                    "AX", "ロフラゼプ酸エチル1mg", "先発", 100, "1,2",  10)
D("メイラックス錠2mg",                    "AX", "ロフラゼプ酸エチル2mg", "先発", 100, "1",    5)
D("ロフラゼプ酸エチル錠1mg「サワイ」",     "AX", "ロフラゼプ酸エチル1mg", "後発", 100, "1,2",  6)
D("クロナゼパム錠0.5mg「アメル」",         "AX", "クロナゼパム0.5mg",    "後発", 100, "1,2,3", 8)
D("リボトリール錠0.5mg",                  "AX", "クロナゼパム0.5mg",    "先発", 100, "1,2,3", 4)
D("リボトリール錠1mg",                    "AX", "クロナゼパム1mg",      "先発", 100, "1,2,3", 4)
D("タンドスピロンクエン酸塩錠10mg「サワイ」", "AX", "タンドスピロン10mg", "後発", 100, "3",    8)
D("タンドスピロンクエン酸塩錠5mg「サワイ」",  "AX", "タンドスピロン5mg",  "後発", 100, "3",    3)
D("アタラックス-Pカプセル25mg",           "AX", "ヒドロキシジン25mg",   "先発", 100, "1,2,3", 4, unit="cap")
# ---- 睡眠薬 ----
D("ゾルピデム酒石酸塩錠5mg「サワイ」",     "HY", "ゾルピデム5mg",        "後発", 100, "1",     12, limit=30)
D("ゾルピデム酒石酸塩錠10mg「サワイ」",    "HY", "ゾルピデム10mg",       "後発", 100, "1",     10, limit=30)
D("マイスリー錠5mg",                      "HY", "ゾルピデム5mg",        "先発", 100, "1",     3, limit=30)
D("エスゾピクロン錠1mg「トーワ」",         "HY", "エスゾピクロン1mg",    "後発", 100, "1",     6)
D("エスゾピクロン錠2mg「トーワ」",         "HY", "エスゾピクロン2mg",    "後発", 100, "1",     12)
D("エスゾピクロン錠3mg「トーワ」",         "HY", "エスゾピクロン3mg",    "後発", 100, "1",     5)
D("ルネスタ錠2mg",                        "HY", "エスゾピクロン2mg",    "先発", 100, "1",     4)
D("ゾピクロン錠7.5mg「トーワ」",           "HY", "ゾピクロン7.5mg",      "後発", 100, "1",     5)
D("ブロチゾラム錠0.25mg「サワイ」",        "HY", "ブロチゾラム0.25mg",   "後発", 100, "1",     16, limit=30)
D("レンドルミン錠0.25mg",                 "HY", "ブロチゾラム0.25mg",   "先発", 100, "1",     4, limit=30)
D("フルニトラゼパム錠1mg「アメル」",       "HY", "フルニトラゼパム1mg",  "後発", 100, "1",     8, limit=30)
D("フルニトラゼパム錠2mg「アメル」",       "HY", "フルニトラゼパム2mg",  "後発", 100, "1",     6, limit=30)
D("サイレース錠1mg",                      "HY", "フルニトラゼパム1mg",  "先発", 100, "1",     3, limit=30)
D("トリアゾラム錠0.25mg「日医工」",        "HY", "トリアゾラム0.25mg",   "後発", 100, "1",     6, limit=30)
D("ハルシオン0.25mg錠",                   "HY", "トリアゾラム0.25mg",   "先発", 100, "1",     2, limit=30)
D("ニトラゼパム錠5mg「トーワ」",           "HY", "ニトラゼパム5mg",      "後発", 100, "1",     4)
D("クアゼパム錠15mg「トーワ」",            "HY", "クアゼパム15mg",       "後発", 100, "1",     4, limit=30)
D("クアゼパム錠20mg「トーワ」",            "HY", "クアゼパム20mg",       "後発", 100, "1",     3, limit=30)
D("ベルソムラ錠10mg",                     "HY", "スボレキサント10mg",   "先発", 100, "1",     4)
D("ベルソムラ錠15mg",                     "HY", "スボレキサント15mg",   "先発", 100, "1",     8)
D("ベルソムラ錠20mg",                     "HY", "スボレキサント20mg",   "先発", 100, "1",     8)
D("デエビゴ錠2.5mg",                      "HY", "レンボレキサント2.5mg", "先発", 100, "1",    10)
D("デエビゴ錠5mg",                        "HY", "レンボレキサント5mg",  "先発", 100, "1",     20)
D("デエビゴ錠10mg",                       "HY", "レンボレキサント10mg", "先発", 100, "1",     8)
D("ロゼレム錠8mg",                        "HY", "ラメルテオン8mg",      "先発", 100, "1",     4)
D("ラメルテオン錠8mg「サワイ」",           "HY", "ラメルテオン8mg",      "後発", 100, "1",     5)
D("クービビック錠25mg",                   "HY", "ダリドレキサント25mg", "先発", 100, "1",     3)
D("クービビック錠50mg",                   "HY", "ダリドレキサント50mg", "先発", 100, "1",     4)
# ---- ADHD ----
D("アトモキセチンカプセル10mg「サワイ」",  "ADHD", "アトモキセチン10mg",  "後発", 100, "1,2",   3, unit="cap")
D("アトモキセチンカプセル25mg「サワイ」",  "ADHD", "アトモキセチン25mg",  "後発", 100, "1,2",   5, unit="cap")
D("アトモキセチンカプセル40mg「サワイ」",  "ADHD", "アトモキセチン40mg",  "後発", 100, "1,2",   6, unit="cap")
D("ストラテラカプセル40mg",               "ADHD", "アトモキセチン40mg",  "先発", 100, "1,2",   2, unit="cap")
D("インチュニブ錠1mg",                    "ADHD", "グアンファシン1mg",   "先発", 100, "1,2",   5)
D("インチュニブ錠3mg",                    "ADHD", "グアンファシン3mg",   "先発", 100, "1,2",   4)
D("コンサータ錠18mg",                     "ADHD", "メチルフェニデート18mg", "先発", 100, "1,2", 6)
D("コンサータ錠27mg",                     "ADHD", "メチルフェニデート27mg", "先発", 100, "1,2", 5)
D("コンサータ錠36mg",                     "ADHD", "メチルフェニデート36mg", "先発", 100, "1,2", 4)
# ---- 認知症 ----
D("ドネペジル塩酸塩錠5mg「サワイ」",       "DEM", "ドネペジル5mg",        "後発", 100, "1",     6)
D("ドネペジル塩酸塩OD錠5mg「トーワ」",     "DEM", "ドネペジル5mg",        "後発", 100, "1",     3)
D("メマンチン塩酸塩錠10mg「DSEP」",        "DEM", "メマンチン10mg",       "後発", 100, "1,2",   4)
D("メマンチン塩酸塩錠20mg「DSEP」",        "DEM", "メマンチン20mg",       "後発", 100, "1",     5)
# ---- 副作用止め・その他 ----
D("アキネトン錠1mg",                      "AC", "ビペリデン1mg",         "先発", 100, "1,2,3", 6)
D("ビペリデン塩酸塩錠1mg「アメル」",       "AC", "ビペリデン1mg",         "後発", 100, "1,2,3", 4)
D("トリヘキシフェニジル塩酸塩錠2mg「アメル」", "AC", "トリヘキシフェニジル2mg", "後発", 100, "1,2,3", 4)
D("ヒベルナ錠25mg",                       "AC", "プロメタジン25mg",      "先発", 100, "1,2,3", 2)
D("プロプラノロール塩酸塩錠10mg「日医工」", "AC", "プロプラノロール10mg", "後発", 100, "1,2,3", 6)
# ---- 内科系（便秘・胃薬など。精神科からも内科からも出る） ----
D("酸化マグネシウム錠330mg「ヨシダ」",     "GI", "酸化マグネシウム330mg", "後発", 500, "2,3,6", 18)
D("センノシド錠12mg「サワイ」",            "GI", "センノシド12mg",        "後発", 500, "1,2",   8)
D("ランソプラゾールOD錠15mg「トーワ」",    "GI", "ランソプラゾール15mg",  "後発", 100, "1",     8)
D("ファモチジンD錠20mg「サワイ」",         "GI", "ファモチジン20mg",      "後発", 100, "1,2",   4)
D("モサプリドクエン酸塩錠5mg「サワイ」",   "GI", "モサプリド5mg",         "後発", 100, "3",     4)
D("ロキソプロフェンNa錠60mg「サワイ」",    "GI", "ロキソプロフェン60mg",  "後発", 100, "1,2,3", 4)
D("アムロジピン錠5mg「サワイ」",           "GI", "アムロジピン5mg",       "後発", 100, "1",     5)
D("ビオフェルミン錠剤",                   "GI", "ビフィズス菌",          "先発", 100, "3,6",   3)
# ---- 出番の少ない薬（年に数回～数十回。間欠・不動の検証用） ----
D("セパゾン錠1mg",                        "AX", "クロキサゾラム1mg",     "先発", 100, "1,2,3", 0.2)
D("レキソタン錠2mg",                      "AX", "ブロマゼパム2mg",       "先発", 100, "1,2,3", 0.3)
D("リーゼ錠5mg",                          "AX", "クロチアゼパム5mg",     "先発", 100, "1,2,3", 0.3)
D("セディール錠10mg",                     "AX", "タンドスピロン10mg",    "先発", 100, "3",     0.3)
D("コンスタン錠0.4mg",                    "AX", "アルプラゾラム0.4mg",   "先発", 100, "1,2,3", 0.2)
D("ドグマチール錠50mg",                   "AD", "スルピリド50mg",        "先発", 100, "1,2,3", 0.2)
D("ルボックス錠25mg",                     "AD", "フルボキサミン25mg",    "先発", 100, "2",     0.2)
D("ミルナシプラン塩酸塩錠25mg「サワイ」",  "AD", "ミルナシプラン25mg",    "後発", 100, "2,4",   0.4)
D("トレドミン錠25mg",                     "AD", "ミルナシプラン25mg",    "先発", 100, "2,4",   0.2)
D("サインバルタカプセル20mg",             "AD", "デュロキセチン20mg",    "先発", 100, "1,2",   0.3, unit="cap")
D("レスリン錠25mg",                       "AD", "トラゾドン25mg",        "先発", 100, "1,2,4", 0.2)
D("ノリトレン錠25mg",                     "AD", "ノルトリプチリン25mg",  "先発", 100, "1,2,3", 0.2)
D("トフラニール錠25mg",                   "AD", "イミプラミン25mg",      "先発", 100, "1,2,3", 0.2)
D("ルジオミール錠25mg",                   "AD", "マプロチリン25mg",      "先発", 100, "1,2",   0.1)
D("ジプレキサ錠5mg",                      "AP", "オランザピン5mg",       "先発", 100, "1,2",   0.3)
D("セロクエル錠25mg",                     "AP", "クエチアピン25mg",      "先発", 100, "1,2,3", 0.3)
D("リスパダール錠1mg",                    "AP", "リスペリドン1mg",       "先発", 100, "1,2",   0.2)
D("セレネース錠1mg",                      "AP", "ハロペリドール1mg",     "先発", 100, "1,2,3", 0.3)
D("エビリファイOD錠6mg",                  "AP", "アリピプラゾール6mg",   "先発", 100, "1,2",   0.2)
D("フルメジン糖衣錠1mg",                  "AP", "フルフェナジン1mg",     "先発", 100, "1,2,3", 0.1)
D("ニューレプチル錠5mg",                  "AP", "プロペリシアジン5mg",   "先発", 100, "1,2,3", 0.1)
D("ピーゼットシー糖衣錠4mg",              "AP", "ペルフェナジン4mg",     "先発", 100, "1,2,3", 0.1)
D("テグレトール錠200mg",                  "MS", "カルバマゼピン200mg",   "先発", 100, "2,3",   0.3)
D("セレニカR錠200mg",                     "MS", "バルプロ酸徐放200mg",   "先発", 100, "2,3,4", 0.3)
D("ラミクタール錠100mg",                  "MS", "ラモトリギン100mg",     "先発", 100, "1,2",   0.3)
D("ガバペン錠200mg",                      "MS", "ガバペンチン200mg",     "先発", 100, "3,6",   0.2)
D("アモバン錠7.5mg",                      "HY", "ゾピクロン7.5mg",       "先発", 100, "1",     0.3)
D("ベンザリン錠5mg",                      "HY", "ニトラゼパム5mg",       "先発", 100, "1",     0.3)
D("ユーロジン2mg錠",                      "HY", "エスタゾラム2mg",       "先発", 100, "1",     0.3)
D("エスタゾラム錠2mg「アメル」",           "HY", "エスタゾラム2mg",       "後発", 100, "1",     0.3)
D("リスミー錠2mg",                        "HY", "リルマザホン2mg",       "先発", 100, "1",     0.3)
D("ダルメートカプセル15mg",               "HY", "フルラゼパム15mg",      "先発", 100, "1",     0.1, unit="cap")
D("エリミン錠5mg",                        "HY", "ニメタゼパム5mg",       "先発", 100, "1",     0.1)
D("アリセプトD錠5mg",                     "DEM", "ドネペジル5mg",        "先発", 100, "1",     0.3)
D("メマリー錠20mg",                       "DEM", "メマンチン20mg",       "先発", 100, "1",     0.3)
D("レミニール錠4mg",                      "DEM", "ガランタミン4mg",      "先発", 100, "2",     0.2)
D("ストラテラカプセル25mg",               "ADHD", "アトモキセチン25mg",  "先発", 100, "1,2",   0.2, unit="cap")
D("ビバンセカプセル30mg",                 "ADHD", "リスデキサンフェタミン30mg", "先発", 100, "1", 0.2, unit="cap")
D("アーテン錠2mg",                        "AC", "トリヘキシフェニジル2mg", "先発", 100, "1,2,3", 0.2)
D("ピレチア錠25mg",                       "AC", "プロメタジン25mg",      "先発", 100, "1,2,3", 0.2)
D("インデラル錠10mg",                     "AC", "プロプラノロール10mg",  "先発", 100, "1,2,3", 0.2)
# ---- もう出ていない薬（在庫だけ残っている。出納にも予定にも出ない） ----
D("ワイパックス錠1mg",                    "AX", "ロラゼパム1mg",         "先発", 100, "1,2,3", 0)
D("セルシン錠5mg",                        "AX", "ジアゼパム5mg",         "先発", 100, "1,2,3", 0)
D("デジレル錠25mg",                       "AD", "トラゾドン25mg",        "先発", 100, "1,2,4", 0)
D("レンドルミンD錠0.25mg",                "HY", "ブロチゾラム0.25mg",    "先発", 100, "1",     0, limit=30)
D("ロヒプノール錠1mg",                    "HY", "フルニトラゼパム1mg",   "先発", 100, "1",     0, limit=30)

NAME = {d["name"]: d for d in DRUGS}
# よく出る後発品は 500 錠包装
BIG_PACK = ("セルトラリン錠25mg「アメル」", "セルトラリン錠50mg「アメル」", "エスシタロプラム錠10mg「DSEP」", "エチゾラム錠0.5mg「トーワ」",
            "ロラゼパム錠0.5mg「サワイ」", "アルプラゾラム錠0.4mg「トーワ」", "ゾルピデム酒石酸塩錠5mg「サワイ」", "ゾルピデム酒石酸塩錠10mg「サワイ」",
            "ブロチゾラム錠0.25mg「サワイ」", "アリピプラゾール錠3mg「明治」", "クエチアピン錠25mg「アメル」", "ミルタザピン錠15mg「サワイ」",
            "トラゾドン塩酸塩錠25mg「アメル」", "エスゾピクロン錠2mg「トーワ」", "オランザピン錠5mg「DSEP」", "リスペリドン錠1mg「アメル」")
for _n in BIG_PACK:
    NAME[_n]["pack"] = 500
# 補助の薬（ほかの薬の副作用止めなど）は単独では出ない：その前提の重みは上の w に含めた

# ------------------------------------------------------------------------------
# 略称（半角カナ）：エリシアS の略称欄に似せる
# ------------------------------------------------------------------------------
_FW = ("ァアィイゥウェエォオカガキギクグケゲコゴサザシジスズセゼソゾタダチヂッツヅテデトド"
       "ナニヌネノハバパヒビピフブプヘベペホボポマミムメモャヤュユョヨラリルレロワヲンヴー")
_HW = ("ｧ ｱ ｨ ｲ ｩ ｳ ｪ ｴ ｫ ｵ ｶ ｶﾞ ｷ ｷﾞ ｸ ｸﾞ ｹ ｹﾞ ｺ ｺﾞ ｻ ｻﾞ ｼ ｼﾞ ｽ ｽﾞ ｾ ｾﾞ ｿ ｿﾞ ﾀ ﾀﾞ ﾁ ﾁﾞ ｯ ﾂ ﾂﾞ ﾃ ﾃﾞ ﾄ ﾄﾞ "
       "ﾅ ﾆ ﾇ ﾈ ﾉ ﾊ ﾊﾞ ﾊﾟ ﾋ ﾋﾞ ﾋﾟ ﾌ ﾌﾞ ﾌﾟ ﾍ ﾍﾞ ﾍﾟ ﾎ ﾎﾞ ﾎﾟ ﾏ ﾐ ﾑ ﾒ ﾓ ｬ ﾔ ｭ ﾕ ｮ ﾖ ﾗ ﾘ ﾙ ﾚ ﾛ ﾜ ｦ ﾝ ｳﾞ ｰ").split(" ")
assert len(_FW) == len(_HW)
_K2H = dict(zip(_FW, _HW))


_PLAIN = str.maketrans("ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポァィゥェォャュョッヴ",
                       "カキクケコサシスセソタチツテトハヒフヘホハヒフヘホアイウエオヤユヨツウ")
_ABBR = {}


def abbr(name):
    """略称：薬品名の先頭3文字のカナ（濁点・半濁点なし、小さい字は大きく）＋2桁の番号。半角カナ"""
    if name in _ABBR:
        return _ABBR[name]
    kana = "".join(c for c in name.translate(_PLAIN) if "ア" <= c <= "ン")[:3]
    used = {v for k, v in _ABBR.items()}
    num = (sum(ord(c) for c in name) * 7) % 99 + 1
    while f"{kana}{num:02d}" in used:
        num = num % 99 + 1
    _ABBR[name] = "".join(_K2H.get(c, c) for c in kana) + f"{num:02d}"
    return _ABBR[name]


# ------------------------------------------------------------------------------
# 患者の処方パターン
# ------------------------------------------------------------------------------
DIAG = [("うつ病", 32), ("不安障害", 18), ("統合失調症", 14), ("双極性障害", 9), ("不眠症", 14),
        ("発達障害", 6), ("認知症", 4), ("適応障害ほか", 3)]
# 診断群ごとの「薬の種類と出る確率」（種類, 確率, 最大剤数）
TEMPLATE = {
    "うつ病":       [("AD", 1.0, 1), ("AD", 0.15, 1), ("AP", 0.15, 1), ("HY", 0.6, 1), ("AX", 0.5, 1), ("GI", 0.15, 1)],
    "不安障害":     [("AD", 0.8, 1), ("AX", 0.85, 1), ("AX", 0.2, 1), ("HY", 0.35, 1), ("GI", 0.1, 1)],
    "統合失調症":   [("AP", 1.0, 1), ("AP", 0.3, 1), ("AC", 0.35, 1), ("HY", 0.5, 1), ("AX", 0.3, 1), ("GI", 0.3, 1), ("MS", 0.1, 1)],
    "双極性障害":   [("MS", 1.0, 1), ("AP", 0.5, 1), ("HY", 0.5, 1), ("AX", 0.2, 1), ("AD", 0.1, 1), ("GI", 0.1, 1)],
    "不眠症":       [("HY", 1.0, 1), ("HY", 0.2, 1), ("AX", 0.2, 1), ("AD", 0.1, 1)],
    "発達障害":     [("ADHD", 1.0, 1), ("HY", 0.2, 1), ("AD", 0.2, 1), ("AX", 0.1, 1)],
    "認知症":       [("DEM", 1.0, 1), ("AP", 0.3, 1), ("HY", 0.3, 1), ("GI", 0.3, 1)],
    "適応障害ほか": [("AX", 0.7, 1), ("AD", 0.4, 1), ("HY", 0.4, 1)],
}
# 診断群ごとの来局間隔（日）と重み
INTERVAL = {
    "うつ病":       [(14, 20), (28, 55), (42, 10), (56, 15)],
    "不安障害":     [(14, 15), (28, 55), (42, 10), (56, 20)],
    "統合失調症":   [(14, 25), (28, 50), (56, 15), (84, 10)],
    "双極性障害":   [(14, 15), (28, 55), (56, 20), (84, 10)],
    "不眠症":       [(14, 15), (28, 55), (30, 15), (56, 15)],
    "発達障害":     [(14, 10), (28, 60), (56, 30)],
    "認知症":       [(28, 50), (56, 30), (84, 20)],
    "適応障害ほか": [(14, 40), (28, 50), (42, 10)],
}
USAGE = {1: "1日1回", 2: "1日2回", 3: "1日3回", 4: "1日2回", 6: "1日3回"}
TIMING = {"HY": "就寝前", "AX": "朝食後", "AD": "朝食後", "AP": "夕食後", "MS": "夕食後", "ADHD": "朝食後",
          "DEM": "朝食後", "AC": "朝食後", "GI": "朝食後"}


def pick(rnd, pairs):
    """(値, 重み) の列から重み付きで1つ選ぶ"""
    vals, ws = zip(*pairs)
    return rnd.choices(vals, weights=ws, k=1)[0]


def pick_drug(rnd, cls, exclude):
    pool = [d for d in DRUGS if d["cls"] == cls and d["group"] not in exclude and d["w"] > 0]
    return rnd.choices(pool, weights=[d["w"] for d in pool], k=1)[0]


def make_patient(rnd, pid):
    """患者1人：診断群・来局間隔・処方内容（薬と1日量）"""
    diag = pick(rnd, DIAG)
    regimen = []
    used = set()
    for cls, p, _ in TEMPLATE[diag]:
        if rnd.random() < p:
            d = pick_drug(rnd, cls, used)
            used.add(d["group"])
            dose = rnd.choice(d["doses"])
            regimen.append((d["name"], dose))
    if not regimen:   # 最低1剤
        d = pick_drug(rnd, TEMPLATE[diag][0][0], used)
        regimen.append((d["name"], rnd.choice(d["doses"])))
    interval = pick(rnd, INTERVAL[diag])
    return dict(id=pid, diag=diag, interval=interval, regimen=regimen)


def rx_days(pt, kind, rnd):
    """処方日数：来局間隔を基本に、30 日制限の薬があれば 28 日まで。臨時は短い"""
    if kind == "臨時":
        return rnd.choice([7, 14])
    days = pt["interval"]
    if kind == "新患":
        days = min(days, rnd.choice([7, 14, 14, 28]))
    limit = min((NAME[n]["limit"] for n, _ in pt["regimen"] if NAME[n]["limit"]), default=None)
    if limit and days > 28:
        days = 28
    return days


BEDTIME = ("ミルタザピン", "リフレックス", "トラゾドン", "テトラミド", "アミトリプチリン", "クエチアピン",
           "ヒルナミン", "コントミン", "オランザピン", "ロドピン")   # 眠気の出る薬は就寝前にまとめる


def usage_text(name, dose):
    d = NAME[name]
    times = USAGE.get(int(dose), "1日1回") if dose >= 1 else "1日1回"
    if d["cls"] == "HY" or (times == "1日1回" and name.startswith(BEDTIME)):
        return "1日1回 就寝前"
    if times == "1日1回":
        return f"1日1回 {TIMING[d['cls']]}"
    if times == "1日2回":
        return "1日2回 朝夕食後"
    return "1日3回 毎食後"

# ------------------------------------------------------------------------------
# 営業日：定休日なし（日曜・祝日も開ける）。年末年始（12/31～1/3）だけ休み
# ------------------------------------------------------------------------------
HOLIDAYS = {dt.date(*x) for x in [
    (2025, 10, 13), (2025, 11, 3), (2025, 11, 24), (2025, 12, 30), (2025, 12, 31),
    (2026, 1, 1), (2026, 1, 2), (2026, 1, 3), (2026, 1, 12), (2026, 2, 11), (2026, 2, 23), (2026, 3, 20),
    (2026, 4, 29), (2026, 5, 4), (2026, 5, 5), (2026, 5, 6), (2026, 7, 20), (2026, 8, 11),
    (2026, 9, 21), (2026, 9, 22), (2026, 9, 23), (2026, 10, 12), (2026, 11, 3), (2026, 11, 23),
    (2026, 12, 30), (2026, 12, 31), (2027, 1, 1), (2027, 1, 2), (2027, 1, 3)]}
WD_WEIGHT = {0: 1.15, 1: 1.0, 2: 0.95, 3: 0.85, 4: 1.0, 5: 1.1, 6: 0.7}   # 曜日の混み具合（月・土が多め、日曜は少なめ）


def is_open(d):
    return not (d.month == 12 and d.day == 31 or d.month == 1 and d.day <= 3)   # 定休日なし（年末年始だけ休み）


def book(p, day, days, rnd):
    """次回予約：day＋days に一番近い「いつもの曜日」（±3日）。休みなら前後の営業日（曜日の好みは残す）"""
    target = day + dt.timedelta(days=days)
    delta = (p["wd"] - target.weekday()) % 7
    if delta > 3:
        delta -= 7
    cand = target + dt.timedelta(days=delta)
    if cand <= day:
        cand += dt.timedelta(days=7)
    return cand if is_open(cand) else next_open(cand, rnd)


def next_open(d, rnd=None):
    """休診日なら前後の営業日へ（予約の取り直し。rnd があれば前日側にもずらす）"""
    if is_open(d):
        return d
    back = d
    while not is_open(back):
        back -= dt.timedelta(days=1)
    fwd = d
    while not is_open(fwd):
        fwd += dt.timedelta(days=1)
    return back if rnd is not None and rnd.random() < 0.5 else fwd


# ------------------------------------------------------------------------------
# 1年分の流れ：患者を固定して日ごとに進める
#   ・定期患者は予約日（前回＋間隔。休診なら次の営業日）にきちんと来る
#   ・来局時のできごと：処方終了（通院をやめる）／用量変更／薬の追加・中止／間隔の変更／残薬調整
#   ・新患が毎日少し入る（初回は 7 or 14 日、その後は定期の間隔）
#   ・臨時（予約外の短い処方）が少し入る。予定（分析データ）には出ない
#   ・内科の処方箋を毎回持ってくる患者が少数いる
# ------------------------------------------------------------------------------
P_STOP, P_DOSE, P_ADD, P_DROP, P_INTERVAL, P_ADJUST = 0.025, 0.04, 0.03, 0.03, 0.03, 0.08
P_NAIKA, P_RINJI = 0.04, 0.015
# 途中で起きる切替：その日以降の来局で、処方が別の薬に置き換わる
#   ・出荷調整：セルトラリン25「アメル」が入らなくなり「トーワ」へ（2026/4 から）
#   ・経過措置：テトラミドが終了し、トラゾドンへ（2026/3 から）
SWITCH = {"セルトラリン錠25mg「アメル」": ("セルトラリン錠25mg「トーワ」", dt.date(2026, 4, 1)),
          "テトラミド錠10mg":            ("トラゾドン塩酸塩錠25mg「アメル」", dt.date(2026, 3, 1))}


def init_population(rnd, start, per_day):
    """開始日時点の通院患者。人数は 1日の処方箋枚数の目安から決める（≒ 平均間隔 30 日・営業日 6/7）"""
    pts, rate = [], 0.0                      # rate：1日あたりの来局の期待値（Σ 1/間隔）
    while rate < per_day:
        p = make_patient(rnd, f"P{10000 + len(pts)}")
        rate += 1 / p["interval"]
        p["naika"] = rnd.random() < P_NAIKA
        p["visits"] = rnd.randint(2, 40)
        # 次回予約：開始日から間隔内の営業日のどこか。曜日の混み具合で偏らせる
        cands = [start + dt.timedelta(days=k) for k in range(p["interval"]) if is_open(start + dt.timedelta(days=k))]
        p["next"] = rnd.choices(cands, weights=[WD_WEIGHT[c.weekday()] for c in cands], k=1)[0]
        p["wd"] = p["next"].weekday()          # いつもの曜日
        days = rx_days(p, "定期", rnd)         # 次回の予定（前回と同じ薬・同じ日数で立っている）
        p["plan"] = [(n, q * days) for n, q in p["regimen"]]
        p["last_lines"] = None
        pts.append(p)
    return pts


def visit(rnd, p, day, kind):
    """1回の来局：できごとを起こし、処方箋（1～2枚）を返す。next を更新する"""
    reg = p["regimen"]
    note = []
    for i, (name, dose) in enumerate(reg):              # 切替（出荷調整・経過措置）
        if name in SWITCH and day >= SWITCH[name][1]:
            new = SWITCH[name][0]
            if new not in [n for n, _ in reg]:
                reg[i] = (new, dose if dose in NAME[new]["doses"] else NAME[new]["doses"][0]); note.append("切替")
            else:
                reg.pop(i); note.append("中止")
            break
    if kind == "定期":
        if rnd.random() < P_DOSE:
            i = rnd.randrange(len(reg)); name, dose = reg[i]
            opts = [x for x in NAME[name]["doses"] if x != dose]
            if opts:
                reg[i] = (name, rnd.choice(opts)); note.append("用量変更")
        if rnd.random() < P_ADD:
            cls = rnd.choice([c for c, _, _ in TEMPLATE[p["diag"]]])
            used = {NAME[n]["group"] for n, _ in reg}
            d = pick_drug(rnd, cls, used)
            if d["group"] not in used:
                reg.append((d["name"], rnd.choice(d["doses"]))); note.append("追加")
        if rnd.random() < P_DROP and len(reg) >= 2:
            reg.pop(rnd.randrange(len(reg))); note.append("中止")
        if rnd.random() < P_INTERVAL:
            p["interval"] = pick(rnd, INTERVAL[p["diag"]])
    days = rx_days(p, kind, rnd)
    adj = rnd.random() < P_ADJUST and kind == "定期" and len(reg) >= 2
    adj_idx = rnd.randrange(len(reg)) if adj else -1
    lines = []
    for j, (name, dose) in enumerate(reg):
        dd = max(7, days - rnd.choice([7, 14])) if j == adj_idx else days
        lines.append(dict(薬品名=name, 用法=usage_text(name, dose), 一日量=dose, 日数=dd, 数量=dose * dd,
                          備考="残薬調整" if j == adj_idx else ""))
    sheets = [dict(患者ID=p["id"], 科="精神科", 区分=kind, 診断群=p["diag"], 日数=days, lines=lines,
                   備考="・".join(note))]
    if p["naika"] and kind != "臨時":
        if "naika_reg" not in p:
            used = set(); regs = []
            for _ in range(rnd.choice([1, 1, 2])):
                d = pick_drug(rnd, "GI", used); used.add(d["group"])
                regs.append((d["name"], rnd.choice(d["doses"])))
            p["naika_reg"] = regs
        nd = min(days, 28) if days <= 28 else days
        sheets.append(dict(患者ID=p["id"], 科="内科", 区分=kind, 診断群=p["diag"], 日数=nd, 備考="",
                           lines=[dict(薬品名=n, 用法=usage_text(n, q), 一日量=q, 日数=nd, 数量=q * nd, 備考="")
                                  for n, q in p["naika_reg"]]))
    if kind != "臨時":
        p["visits"] += 1
        # 予定（分析データ）は「次回予約日に、今回と同じ薬・同じ日数」で立つ
        p["plan"] = [(n, q * days) for n, q in reg]
        if kind == "新患" and days < p["interval"]:
            p["next"] = book(p, day, days, rnd)
        elif rnd.random() < P_STOP and p["visits"] > 2:
            p["next"] = None; p["plan"] = None          # 通院終了（次の予約を取らない）
        else:   # 次回は処方日数のあと（30日制限で 28 日になった人は 28 日後）
            p["next"] = book(p, day, days, rnd)
        if p["next"] and kind != "新患":
            p["plan_days"] = days
    return sheets


def simulate(start, end, per_day):
    rnd = random.Random(SEED)
    pts = init_population(rnd, start, per_day)
    new_rate = per_day * P_STOP                   # 新患／営業日（やめる人数と釣り合わせる）
    sched = defaultdict(list)
    for p in pts:
        sched[p["next"]].append(p)
    pid_next = 10000 + len(pts)
    visits, plan = [], defaultdict(dict)          # plan[day][name] = (数量, 患者数)
    d = start
    while d <= end:
        if is_open(d):
            today = sched.pop(d, [])
            rnd.shuffle(today)
            for p in today:
                # 予定（来局前の時点の予約内容）
                for n, q in (p.get("plan") or []):
                    s = plan[d].get(n, (0, 0)); plan[d][n] = (s[0] + q, s[1] + 1)
                kind = "定期"
                for s in visit(rnd, p, d, kind):
                    visits.append((d, s))
                if p["next"]:
                    sched[p["next"]].append(p)
            # 新患
            for _ in range(poisson(rnd, new_rate)):
                p = make_patient(rnd, f"P{pid_next}"); pid_next += 1
                p["naika"] = rnd.random() < P_NAIKA; p["visits"] = 0; p["plan"] = None; p["wd"] = d.weekday()
                for s in visit(rnd, p, d, "新患"):
                    visits.append((d, s))
                if p["next"]:
                    sched[p["next"]].append(p); pts.append(p)
            # 臨時（予約外）：通院中の患者が短い処方を取りに来る
            active = [p for p in pts if p["next"]]
            for _ in range(poisson(rnd, len(today) * P_RINJI)):
                p = rnd.choice(active)
                for s in visit(rnd, p, d, "臨時"):
                    s["lines"] = s["lines"][:1]        # 1剤だけ
                    visits.append((d, s))
        d += dt.timedelta(days=1)
    # 未来の予定（終了日以降に予約がある患者）も分析データには載るが、ここでは期間内だけ扱う
    return visits, plan


def poisson(rnd, lam):
    if lam <= 0:
        return 0
    L, k, p = math.exp(-lam), 0, 1.0
    while True:
        p *= rnd.random()
        if p < L:
            return k
        k += 1


def fmt_num(x):
    return int(x) if float(x).is_integer() else x


CLS_JA = dict(AD="抗うつ薬", AP="抗精神病薬", MS="気分安定薬", AX="抗不安薬", HY="睡眠薬", ADHD="ADHD治療薬",
              DEM="認知症治療薬", AC="副作用止めほか", GI="内科系")


def write_period(start, end, visits, plan):
    folder = os.path.join(OUT, f"{start:%Y%m%d}_{end:%Y%m%d}")
    os.makedirs(folder, exist_ok=True)
    HDR = ["処方日", "処方箋No", "患者ID", "科", "来局区分", "診断群", "処方日数", "薬品名", "略称", "用法",
           "1日量", "日数", "数量", "単位", "成分規格", "需要のまとまり", "先発後発", "備考"]
    rows, sheet_rows = [], []
    for no, (d, s) in enumerate(visits, 1):
        sheet_rows.append([d.strftime("%Y/%m/%d"), no, s["患者ID"], s["科"], s["区分"], s["診断群"], s["日数"],
                           len(s["lines"]), s["備考"]])
        for l in s["lines"]:
            dr = NAME[l["薬品名"]]
            rows.append([d.strftime("%Y/%m/%d"), no, s["患者ID"], s["科"], s["区分"], s["診断群"], s["日数"],
                         l["薬品名"], abbr(l["薬品名"]), l["用法"], fmt_num(l["一日量"]), l["日数"], fmt_num(l["数量"]),
                         dr["unit"], dr["group"], dr["cluster"], dr["gen"], l["備考"]])
    with open(os.path.join(folder, "処方明細.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(HDR); w.writerows(rows)
    # 日別・薬品別：予定と実績
    act = defaultdict(lambda: [0, 0.0])
    for d, s in visits:
        for l in s["lines"]:
            act[(d, l["薬品名"])][0] += 1; act[(d, l["薬品名"])][1] += l["数量"]
    keys = sorted(set(act) | {(d, n) for d in plan for n in plan[d]})
    pa_rows = [[d.strftime("%Y/%m/%d"), n, plan[d].get(n, (0, 0))[1], fmt_num(plan[d].get(n, (0, 0))[0]),
                act.get((d, n), [0, 0])[0], fmt_num(act.get((d, n), [0, 0])[1]),
                fmt_num(act.get((d, n), [0, 0])[1] - plan[d].get(n, (0, 0))[0])] for d, n in keys]
    with open(os.path.join(folder, "日別_予定と実績.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f); w.writerow(["日付", "薬品名", "予定患者数", "予定数量", "実績患者数", "実績数量", "ずれ"]); w.writerows(pa_rows)
    # ブック（明細が多いときは明細シートを省く）
    wb = Workbook(); wb.remove(wb.active)
    if len(rows) <= 60000:
        ws = wb.create_sheet("処方明細"); ws.append(HDR)
        for r in rows:
            ws.append(r)
    ws = wb.create_sheet("処方箋一覧"); ws.append(["処方日", "処方箋No", "患者ID", "科", "来局区分", "診断群", "処方日数", "薬剤数", "できごと"])
    for r in sheet_rows:
        ws.append(r)
    ws = wb.create_sheet("日別集計"); ws.append(["日付", "曜日", "処方箋", "患者数", "定期", "新患", "臨時", "内科", "明細行"])
    by_day = defaultdict(list)
    for d, s in visits:
        by_day[d].append(s)
    for d in sorted(by_day):
        ss = by_day[d]
        ws.append([d.strftime("%Y/%m/%d"), WD[d.weekday()], len(ss), len({s["患者ID"] for s in ss}),
                   sum(s["区分"] == "定期" for s in ss), sum(s["区分"] == "新患" for s in ss),
                   sum(s["区分"] == "臨時" for s in ss), sum(s["科"] == "内科" for s in ss), sum(len(s["lines"]) for s in ss)])
    agg = defaultdict(lambda: [0, 0.0])
    for d, s in visits:
        for l in s["lines"]:
            agg[l["薬品名"]][0] += 1; agg[l["薬品名"]][1] += l["数量"]
    ws = wb.create_sheet("薬品別集計"); ws.append(["製剤", "略称", "薬品名", "分類", "延べ患者数", "数量", "単位", "成分規格", "需要のまとまり", "先発後発"])
    for n, (p, q) in sorted(agg.items(), key=lambda kv: -kv[1][1]):
        dr = NAME[n]
        ws.append(["内服", abbr(n), n, CLS_JA[dr["cls"]], p, fmt_num(q), dr["unit"], dr["group"], dr["cluster"], dr["gen"]])
    if len(pa_rows) <= 200000:
        ws = wb.create_sheet("日別_予定と実績"); ws.append(["日付", "薬品名", "予定患者数", "予定数量", "実績患者数", "実績数量", "ずれ"])
        for r in pa_rows:
            ws.append(r)
    ws = wb.create_sheet("薬品マスタ"); ws.append(["薬品名", "略称", "分類", "成分規格", "需要のまとまり", "先発後発", "包装単位", "単位", "1日量の候補", "日数制限"])
    for dr in DRUGS:
        ws.append([dr["name"], abbr(dr["name"]), CLS_JA[dr["cls"]], dr["group"], dr["cluster"], dr["gen"], dr["pack"], dr["unit"],
                   "・".join(str(fmt_num(x)) for x in dr["doses"]), dr["limit"]])
    for w in wb:
        for c in w[1]:
            c.font = Font(bold=True)
    path = os.path.join(folder, f"処方_{start:%Y%m%d}_{end:%Y%m%d}.xlsx")
    wb.save(path)
    return folder, by_day, agg


def summary(start, end, visits, by_day, agg):
    n = [len(v) for v in by_day.values()]
    print(f"期間：{start:%Y/%m/%d}～{end:%Y/%m/%d}　営業日 {len(by_day)} 日　処方箋 {sum(n)} 枚（1日 {min(n)}～{max(n)}、平均 {sum(n)/len(n):.0f}）　"
          f"患者 {len({s['患者ID'] for _, s in visits})} 人　薬品 {len(agg)} 品目")
    print("来局区分：", dict(Counter(s["区分"] for _, s in visits)), "　科：", dict(Counter(s["科"] for _, s in visits)))
    print("薬剤数/枚：", dict(sorted(Counter(len(s["lines"]) for _, s in visits).items())))
    print("処方日数：", dict(sorted(Counter(s["日数"] for _, s in visits).items())))
    print("できごと：", dict(Counter(x for _, s in visits for x in s["備考"].split("・") if x)))
    def line(d, ss):
        return (f"  {d:%Y/%m/%d}({WD[d.weekday()]}) {len(ss):4d} 枚  定期 {sum(s['区分']=='定期' for s in ss):3d}  新患 {sum(s['区分']=='新患' for s in ss):2d}  "
                f"臨時 {sum(s['区分']=='臨時' for s in ss):2d}  内科 {sum(s['科']=='内科' for s in ss):2d}")
    if len(by_day) <= 31:
        print("日別：")
        for d in sorted(by_day):
            print(line(d, by_day[d]))
    else:
        print("月別：")
        bym = defaultdict(list)
        for d, ss in by_day.items():
            bym[(d.year, d.month)].append(len(ss))
        for (y, mo), v in sorted(bym.items()):
            print(f"  {y}/{mo:02d}  営業日 {len(v):2d}  処方箋 {sum(v):5d}（1日 {min(v)}～{max(v)}、平均 {sum(v)/len(v):.0f}）")
        order = sorted(by_day, key=lambda d: len(by_day[d]))
        print("少ない日："); [print(line(d, by_day[d])) for d in order[:4]]
        print("多い日：");   [print(line(d, by_day[d])) for d in order[-4:]]


# ------------------------------------------------------------------------------
# 在庫の動き（出納の元）：発注点方式で仕入を自動化し、返品・廃棄・移動を少し混ぜる
#   ・卸の配達：月～土（祝日・12/30～1/4 は休み）。発注した翌配達日に入る
#   ・発注点＝1日平均 × 7 日、発注量＝21 日分になるまで包装単位で切り上げ（最低 1 箱）
#   ・足りないときは当日便で補う（欠品させない）
#   ・取り寄せ品（発注除外）は患者ごとに必要量だけ仕入れる
# ------------------------------------------------------------------------------
# 実物（薬品使用量予測データ.xls の構造分析レポート 2026-10-09）どおり：I・J は非表示の列
HDR_PLAN = ["製剤", "略称", "薬品名", "患者数", "予定使用量", "在庫数量", "予定在庫数", "単位",
            "予定使用量（当日までの小計）", "在庫数量（発注残を含まない）"]
HIDDEN_COLS = ("I", "J")
SEIZAI = "内服  "                                   # 実物は「内服」＋末尾空白 2 つの固定長


def disp(name):
    """実物の表記：mg は全角「ｍｇ」（fx_正規化 で半角に戻る）"""
    return name.replace("mg", "ｍｇ")


# 品名の先頭の記号（実物の出納・各リスト：劇薬「劇 」、向精神薬「向 」、なければ空白 3 つ）
MARK = dict(AP="劇", MS="劇", AD="劇", HY="向", AX="向", ADHD="向")


def marked(name):
    m = MARK.get(NAME[name]["cls"])
    return (m + " " if m else "   ") + disp(name)
HDR_SUITOU = ["製剤コード", "製剤", "コード", "品名", "繰越数量", "仕入数量", "払出数量", "在庫数量", "単位", "棚番"]
NO_ORDER = {"クロザリル錠25mg": "取り寄せ（患者ごとに発注）", "クロザリル錠100mg": "取り寄せ（患者ごとに発注）"}
TRANSITIONAL = {"テトラミド錠10mg": dt.date(2026, 3, 31), "パロキセチン錠20mg「日医工」": dt.date(2027, 3, 31)}
SHELF = dict(AD="A", AP="B", MS="C", AX="D", HY="E", ADHD="F", DEM="G", AC="H", GI="J")


def wholesaler_open(d):
    return d.weekday() < 6 and d not in HOLIDAYS and not ((d.month == 12 and d.day >= 30) or (d.month == 1 and d.day <= 4))


def next_delivery(d):
    d += dt.timedelta(days=1)
    while not wholesaler_open(d):
        d += dt.timedelta(days=1)
    return d


def ledger(visits, start, end):
    """日ごとの在庫の動き。戻り値：
       mov[(day, name)] = dict(仕入, 払出, その他)　stock_eod[(day, name)]　pending[(day, name)]（発注残）"""
    rnd = random.Random(SEED + 7)
    dispensed = defaultdict(float)
    for d, s in visits:
        for l in s["lines"]:
            dispensed[(d, l["薬品名"])] += l["数量"]
    ndays = (end - start).days + 1
    avg = defaultdict(float)
    for (d, n), q in dispensed.items():
        avg[n] += q / ndays
    names = [x["name"] for x in DRUGS]
    stock, arrive = {}, defaultdict(lambda: defaultdict(float))
    for n in names:
        pack = NAME[n]["pack"]
        if avg[n] > 0:
            boxes = math.ceil(avg[n] * rnd.uniform(20, 45) / pack)
            stock[n] = boxes * pack - rnd.randrange(0, max(1, min(pack, int(avg[n] * 5) + 1)))
        else:
            stock[n] = pack if rnd.random() < 0.5 else 0     # 使われていない薬：半分は在庫だけ残っている
    last_move = {n: start - dt.timedelta(days=rnd.randint(0, 120)) for n in names}
    mov, stock_eod, pending = {}, {}, {}
    on_order = defaultdict(float)

    def rec(d, n):
        return mov.setdefault((d, n), dict(仕入=0.0, 払出=0.0, その他=0.0))

    d = start
    while d <= end:
        for n in names:
            r = dict(仕入=0.0, 払出=0.0, その他=0.0)
            if arrive[d][n]:                                           # 入荷
                r["仕入"] += arrive[d][n]; stock[n] += arrive[d][n]; on_order[n] -= arrive[d][n]
            q = dispensed.get((d, n), 0.0)                             # 払出
            if q:
                if stock[n] < q:                                       # 当日便で補う
                    pack = NAME[n]["pack"]
                    add = math.ceil((q - stock[n]) / pack) * pack if n not in NO_ORDER else q - stock[n]
                    r["仕入"] += add; stock[n] += add
                r["払出"] += q; stock[n] -= q; last_move[n] = d
            if wholesaler_open(d) and n not in NO_ORDER and avg[n] > 0:   # 発注（卸の営業日の夕方）
                pack = NAME[n]["pack"]
                if stock[n] + on_order[n] < avg[n] * 7:
                    need = avg[n] * 21 - (stock[n] + on_order[n])
                    qty = max(1, math.ceil(need / pack)) * pack
                    arrive[next_delivery(d)][n] += qty; on_order[n] += qty
            if r["仕入"] or r["払出"]:
                mov[(d, n)] = r
        if d.day == 15:   # その他の増減：月に数件（廃棄・返品・店舗間の移動）
            cands = [n for n in names if stock[n] > 0]
            for n in rnd.sample(cands, min(2, len(cands))):           # 期限切れ廃棄
                q = min(stock[n], rnd.randint(5, 30)); stock[n] -= q; rec(d, n)["その他"] -= q
            for n in [n for n in names if stock[n] > 0 and (d - last_move[n]).days >= 90 and avg[n] * 30 < 1]:
                if rnd.random() < 0.3:                                  # 不動品の返品
                    rec(d, n)["その他"] -= stock[n]; stock[n] = 0
            busy = [n for n in names if avg[n] * 30 >= 100]
            if busy:
                n = rnd.choice(busy); pack = NAME[n]["pack"]           # 他店へ譲渡
                if stock[n] > pack:
                    rec(d, n)["その他"] -= pack; stock[n] -= pack
                n = rnd.choice(busy); pack = NAME[n]["pack"]           # 他店から受け入れ
                rec(d, n)["その他"] += pack; stock[n] += pack
        for n in names:
            stock_eod[(d, n)] = stock[n]; pending[(d, n)] = on_order[n]
        d += dt.timedelta(days=1)
    return mov, stock_eod, pending, avg, last_move


def write_xlsx(path, sheets, hidden=()):
    wb = Workbook(); wb.remove(wb.active)
    for title, rows, merged_second in sheets:
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
        if merged_second:   # 実物：1行目が見出し、2行目は空欄で1行目と縦に結合
            for c in range(1, len(rows[0]) + 1):
                ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
        for col in hidden:  # 実物で非表示の列
            ws.column_dimensions[col].hidden = True
        for c in ws[1]:
            c.font = Font(bold=True)
    wb.save(path)


def write_exports(exp_start, exp_end, plan, mov, stock_eod, pending):
    """ツールが読む形式：実績/薬品使用量予測データYYYYMM.xlsx（日別シート）と 出納/<start>_<end>_数量一覧.xlsx"""
    H, S = os.path.join(OUT, "export", "実績"), os.path.join(OUT, "export", "出納")
    os.makedirs(H, exist_ok=True); os.makedirs(S, exist_ok=True)
    days = [exp_start + dt.timedelta(i) for i in range((exp_end - exp_start).days + 1)]
    code = {x["name"]: f"{i + 1:04d}" for i, x in enumerate(DRUGS)}
    shelf, cnt = {}, Counter()
    for x in DRUGS:
        cnt[x["cls"]] += 1; shelf[x["name"]] = f"{SHELF[x['cls']]}-{cnt[x['cls']]:02d}"

    def stock_bod(d, n):   # その日の朝の在庫 ＝ 夜の在庫 － その日の動き
        r = mov.get((d, n)) or dict(仕入=0, 払出=0, その他=0)
        return stock_eod[(d, n)] - r["仕入"] + r["払出"] - r["その他"]

    # ---- 薬品使用量分析データ（予定）：月ファイル。1行目見出し＋2行目空欄（結合）、日ごとのシート ----
    cum = defaultdict(float)                              # I：予定使用量の当日までの小計（月初から）

    def plan_row(n, p, q, d):
        st, pe = stock_eod[(d, n)], pending[(d, n)]       # F 在庫数量は発注残込み、G ＝ F － E、J は発注残なし
        cum[n] += q
        return [SEIZAI, abbr(n), disp(n), p, fmt_num(q), fmt_num(st + pe), fmt_num(st + pe - q), NAME[n]["unit"],
                fmt_num(cum[n]), fmt_num(st)]
    order = lambda n: (abbr(n), n)
    sheets, tot = [], defaultdict(lambda: [0, 0.0, 0.0, 0.0, 0.0])
    for d in days:
        if not is_open(d):
            continue
        pd = plan.get(d, {})
        rows = [HDR_PLAN, [None] * len(HDR_PLAN)]
        for n in sorted(pd, key=order):
            r = plan_row(n, pd[n][1], pd[n][0], d)
            rows.append(r)
            for i, c in enumerate((3, 4, 5, 6, 9)):          # 期間すべて：日別シートの単純な合計（在庫も足されている）
                tot[n][i] += r[c]
        sheets.append((f"{d.month}月{d.day}日({WD[d.weekday()]})", rows, True))
    rows = [HDR_PLAN, [None] * len(HDR_PLAN)]
    for n in sorted(tot, key=order):
        t = tot[n]
        rows.append([SEIZAI, abbr(n), disp(n), int(t[0]), fmt_num(t[1]), fmt_num(t[2]), fmt_num(t[3]), NAME[n]["unit"],
                     0, fmt_num(t[4])])                      # I は期間すべてでは固定値（0）
    sheets.insert(0, ("期間すべて", rows, True))
    plan_path = os.path.join(H, f"薬品使用量予測データ{exp_start:%Y%m}.xlsx")
    write_xlsx(plan_path, sheets, hidden=HIDDEN_COLS)

    # ---- 出納（実際）：期間に動いた薬だけ。繰越＝期間初日の朝、在庫＝期間最終日の夜 ----
    rows = [["数量一覧", None, f"出力日：{exp_end + dt.timedelta(days=1):%Y/%m/%d}"],
            [None, None, f"期間：{exp_start:%Y/%m/%d}～{exp_end:%Y/%m/%d}"],
            [],
            HDR_SUITOU]
    n_rows = 0
    for x in DRUGS:
        n = x["name"]
        agg = dict(仕入=0.0, 払出=0.0, その他=0.0)
        for d in days:
            r = mov.get((d, n))
            if r:
                for k in agg:
                    agg[k] += r[k]
        if not (agg["仕入"] or agg["払出"] or agg["その他"]):
            continue
        rows.append([1, SEIZAI, code[n], marked(n), fmt_num(stock_bod(exp_start, n)), fmt_num(agg["仕入"]), fmt_num(agg["払出"]),
                     fmt_num(stock_eod[(exp_end, n)]), x["unit"], shelf[n]])
        n_rows += 1
    suitou_path = os.path.join(S, f"{exp_start:%Y%m%d}_{exp_end:%Y%m%d}_数量一覧.xlsx")
    write_xlsx(suitou_path, [("Sheet1", rows, False)])
    return plan_path, suitou_path, len(sheets) - 1, n_rows


def write_forecast(fs, fe, asof, plan, stock_eod, pending):
    """期間の翌日からの予測データ（予約から立つ予定）。在庫は書き出し時点（asof の夜）の値。予測フォルダに置く"""
    F = os.path.join(OUT, "export", "予測")
    os.makedirs(F, exist_ok=True)
    days = [fs + dt.timedelta(i) for i in range((fe - fs).days + 1)]
    order = lambda n: (abbr(n), n)
    cum = defaultdict(float)
    sheets, tot = [], defaultdict(lambda: [0, 0.0, 0.0, 0.0, 0.0])
    for d in days:
        pd = plan.get(d, {})
        rows = [HDR_PLAN, [None] * len(HDR_PLAN)]
        for n in sorted(pd, key=order):
            q, p = pd[n]
            st, pe = stock_eod[(asof, n)], pending[(asof, n)]
            cum[n] += q
            r = [SEIZAI, abbr(n), disp(n), p, fmt_num(q), fmt_num(st + pe), fmt_num(st + pe - q), NAME[n]["unit"], fmt_num(cum[n]), fmt_num(st)]
            rows.append(r)
            for i, c in enumerate((3, 4, 5, 6, 9)):
                tot[n][i] += r[c]
        sheets.append((f"{d.month}月{d.day}日({WD[d.weekday()]})", rows, True))
    rows = [HDR_PLAN, [None] * len(HDR_PLAN)]
    for n in sorted(tot, key=order):
        t = tot[n]
        rows.append([SEIZAI, abbr(n), disp(n), int(t[0]), fmt_num(t[1]), fmt_num(t[2]), fmt_num(t[3]), NAME[n]["unit"], 0, fmt_num(t[4])])
    sheets.insert(0, ("期間すべて", rows, True))
    path = os.path.join(F, f"薬品使用量予測データ{fs:%Y%m}.xlsx")
    write_xlsx(path, sheets, hidden=HIDDEN_COLS)
    return path


# ------------------------------------------------------------------------------
# エリシアSの 3 リスト（構造分析レポート 2026-11-03 の列構成どおり）
#   経過措置薬品リスト.xls         ：薬品ごと。品名に記号、経過措置（期限）、在庫数量、単位。記号 × は在庫 0
#   YYYYMMDD_期限切れリスト.xls    ：仕入伝票の行（ロット）。期限・期限切れ数量・ロット番号
#   YYYYMM_不動品リスト.xls        ：仕入伝票の行（ロット）。期限日・最終払出日・不動日数・数量・薬価・YJコード・基準日
#   いずれも見出しは 1 行目、データは 2 行目から、結合なし
# ------------------------------------------------------------------------------
HDR_KEIKA = ["記号", "製剤コード", "製剤", "コード", "品名", "メーカーコード", "メーカー", "経過措置", "在庫数量", "単位"]
HDR_KIGEN = ["記号", "製剤コード", "製剤", "コード", "品名", "メーカーコード", "メーカー", "棚番", "伝票日付", "No", "伝票区分コード", "伝票区分",
             "仕入先コード", "仕入先", "包装", "仕入数量", "単価", "期限", "期限切れ数量", "単位", "期限切れ金額", "ロット番号", "伝票備考",
             "連携ｺｰﾄﾞ", "連番", "品目備考"]
HDR_FUDO = ["記号", "製剤コード", "製剤", "コード", "品名", "メーカーコード", "メーカー", "棚番", "仕入伝票日付", "仕入区分コード", "仕入区分",
            "仕入先コード", "仕入先", "包装", "購入価", "仕入数量", "仕入金額", "期限日", "最終払出日", "払出区分", "払出区分コード", "払出先コード",
            "払出数量", "払出先", "単価", "数量", "単位", "金額", "薬価", "薬価金額", "連携コード", "連番", "品目備考", "ロット番号", "伝票備考",
            "基準日", "最終使用日", "不動日数", "YJコード", "出力日時", "JANコード", "レセプト電算コード"]
WHOLESALERS = [("ｱﾙﾌ", "アルフレッサ"), ("ｽｽﾞｹﾝ", "スズケン"), ("ﾒﾃﾞｨ", "メディセオ")]


def write_system_lists(end, stock_eod, mov, last_move):
    """3 リストを export/システムリスト/<種類>/ に書く。ロットは在庫から架空に分ける"""
    rnd = random.Random(SEED + 9)
    base = end + dt.timedelta(days=1)
    root = os.path.join(OUT, "export", "システムリスト")
    dirs = {k: os.path.join(root, k) for k in ("経過措置薬品", "期限切れ", "不動薬")}
    for d in dirs.values():
        os.makedirs(d, exist_ok=True)
    code = {x["name"]: abbr(x["name"]) for x in DRUGS}
    shelf, cnt = {}, Counter()
    for x in DRUGS:
        cnt[x["cls"]] += 1; shelf[x["name"]] = f"{SHELF[x['cls']]}-{cnt[x['cls']]:02d}"
    price = {x["name"]: round(rnd.uniform(5, 120), 1) if x["gen"] == "後発" else round(rnd.uniform(20, 300), 1) for x in DRUGS}
    yj = {x["name"]: f"{rnd.randint(1124, 1179)}0{rnd.randint(10, 99)}F{rnd.randint(1, 9)}{rnd.randint(100, 999)}" for x in DRUGS}
    jan = {x["name"]: "49" + "".join(str(rnd.randint(0, 9)) for _ in range(11)) for x in DRUGS}
    # 最後に払出があった日
    last_out = {}
    for (d, n), r in mov.items():
        if r["払出"] > 0 and d <= end and (n not in last_out or d > last_out[n]):
            last_out[n] = d
    # ロット：在庫を 1～3 本に分ける
    lots = {}
    for x in DRUGS:
        n = x["name"]; st = stock_eod[(end, n)]
        if st <= 0:
            continue
        k = 1 if st <= x["pack"] else rnd.randint(1, 3)
        cuts = sorted(rnd.sample(range(1, int(st)), k - 1)) if k > 1 and st > k else []
        bounds = [0] + cuts + [int(st)]
        qs = [bounds[i + 1] - bounds[i] for i in range(k)]
        if st != int(st):
            qs[-1] += st - int(st)
        lots[n] = [dict(ロット=f"{rnd.choice('ABCDEFG')}{rnd.randint(1000, 9999)}", 数量=q,
                        期限=end + dt.timedelta(days=rnd.randint(60, 900)),
                        仕入日=end - dt.timedelta(days=rnd.randint(10, 400)),
                        仕入先=rnd.choice(WHOLESALERS)) for q in qs if q > 0]

    def mk(n):
        return [1, "内服", code[n], marked(n), "00000", None]

    # ---- 経過措置薬品リスト ----
    rows = [HDR_KEIKA]
    names = list(TRANSITIONAL) + ["ワイパックス錠1mg", "ロヒプノール錠1mg"]
    for n in names:
        st = stock_eod.get((end, n), 0)
        lim = TRANSITIONAL.get(n, dt.date(2027, 3, 31))
        rows.append(["×" if st <= 0 else None] + mk(n)[:5] + [None, lim.strftime("%Y/%m/%d"), fmt_num(st), NAME[n]["unit"]])
    p1 = os.path.join(dirs["経過措置薬品"], "経過措置薬品リスト.xlsx")
    write_xlsx(p1, [("Sheet1", rows, False)])

    # ---- 期限切れリスト（基準日から 180 日以内に期限が来るロット） ----
    rows = [HDR_KIGEN]
    seq = 0
    for n, ls in lots.items():
        for l in ls:
            if (l["期限"] - base).days <= 180:
                seq += 1
                rows.append([None, 1, "内服", code[n], marked(n), "00000", None, shelf[n], l["仕入日"].strftime("%Y/%m/%d"), 1, 11, "仕入",
                             l["仕入先"][0], l["仕入先"][1], None, fmt_num(l["数量"]), price[n], l["期限"].strftime("%Y/%m/%d"), fmt_num(l["数量"]),
                             NAME[n]["unit"], round(l["数量"] * price[n]), l["ロット"], f"{l['仕入先'][1]}（{price[n]}/{l['ロット']}）",
                             f"5{l['仕入日'].year - 2018:02d}{l['仕入日'].month:02d}", seq, None])
    p2 = os.path.join(dirs["期限切れ"], f"{end:%Y%m%d}_期限切れリスト.xlsx")
    write_xlsx(p2, [("Sheet1", rows, False)])

    # ---- 不動品リスト（基準日の 90 日以上前から払出がない薬のロット） ----
    rows = [HDR_FUDO]
    seq = 0
    for n, ls in lots.items():
        lo = last_out.get(n)
        idle = (base - lo).days if lo else None
        if lo is not None and idle < 90:
            continue
        for l in ls:
            seq += 1
            rows.append([None, 1, "内服", code[n], marked(n), "00000", None, shelf[n], l["仕入日"].strftime("%Y/%m/%d"), 11, "仕入",
                         l["仕入先"][0], l["仕入先"][1], f"{NAME[n]['pack']}{NAME[n]['unit']}", price[n], fmt_num(l["数量"]), round(l["数量"] * price[n]),
                         l["期限"].strftime("%Y/%m/%d"), lo.strftime("%Y/%m/%d") if lo else None, "払出" if lo else None, 21 if lo else None,
                         1 if lo else None, fmt_num(l["数量"]) if lo else None, "処方" if lo else None, price[n], fmt_num(l["数量"]), NAME[n]["unit"],
                         round(l["数量"] * price[n]), price[n], round(l["数量"] * price[n]), f"5{l['仕入日'].year - 2018:02d}{l['仕入日'].month:02d}",
                         seq, None, l["ロット"], f"{l['仕入先'][1]}（{price[n]}/{l['ロット']}）", base.strftime("%Y/%m/%d"),
                         lo.strftime("%Y/%m/%d") if lo else None, idle, yj[n], base.strftime("%Y/%m/%d") + " 09:00:00", jan[n], 6 * 10**8 + seq])
    p3 = os.path.join(dirs["不動薬"], f"{end:%Y%m}_不動品リスト.xlsx")
    write_xlsx(p3, [("Sheet1", rows, False)])
    return [p1, p2, p3]


def write_init_tables(sim_start, sim_end, base_date):
    """分析ブックの手入力テーブルに貼る値（設定・経過措置リスト・発注除外・在庫情報・追加休業日・薬品グループ）"""
    H, S = os.path.join(OUT, "export", "実績"), os.path.join(OUT, "export", "出納")
    wb = Workbook(); wb.remove(wb.active)
    ws = wb.create_sheet("設定"); ws.append(["項目", "値"])
    # 開始日・終了日は書かない（0.3.0：発注日までの 1 年間を自動で読む。発注日 ＝ 期間の翌日なので、期間がそのまま対象になる）
    for k, v in [("フォルダパス", H), ("出納フォルダパス", S), ("年末年始の休み始め", "12/30"), ("年末年始の休み終わり", "1/4"),
                 ("発注日", base_date.strftime("%Y/%m/%d")), ("不動とみなす月数", 3), ("在庫過多とみなす日数", 60),
                 ("ずれの注意割合", 0.3), ("ずれの大きさの注意割合", 0.5),
                 ("期限切迫フォルダパス", os.path.join(OUT, "export", "システムリスト", "期限切れ")),
                 ("経過措置フォルダパス", os.path.join(OUT, "export", "システムリスト", "経過措置薬品")),
                 ("不動品フォルダパス", os.path.join(OUT, "export", "システムリスト", "不動薬"))]:
        ws.append([k, v])
    ws = wb.create_sheet("経過措置リスト"); ws.append(["薬品名", "経過措置期限"])
    for n, d in TRANSITIONAL.items():
        ws.append([n, d.strftime("%Y/%m/%d")])
    ws = wb.create_sheet("発注除外"); ws.append(["薬品名", "理由", "登録日"])
    for n, why in NO_ORDER.items():
        ws.append([n, why, sim_start.strftime("%Y/%m/%d")])
    ws = wb.create_sheet("在庫情報"); ws.append(["薬品名", "列数"])
    for x in DRUGS:
        ws.append([x["name"], 10])
    ws = wb.create_sheet("追加休業日"); ws.append(["日付", "理由"])
    ws = wb.create_sheet("薬品グループ"); ws.append(["薬品名", "成分規格", "需要のまとまり", "先発後発"])
    for x in DRUGS:
        ws.append([x["name"], x["group"], x["cluster"], x["gen"]])
    for w in wb:
        for c in w[1]:
            c.font = Font(bold=True)
    path = os.path.join(OUT, "export", "手入力テーブル.xlsx")
    try:
        wb.save(path)
    except PermissionError:                       # Excel で開いているときは別名に書く
        path = os.path.join(OUT, "export", "手入力テーブル_生成.xlsx")
        wb.save(path)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="start", default="2025-10-01", help="処方を作る期間（既定：1年）")
    ap.add_argument("--to", dest="end", default="2026-09-30")
    ap.add_argument("--per-day", type=int, default=175, help="1日の定期来局の目安（新患・臨時・内科の分が上乗せされる）")
    ap.add_argument("--rx", action="store_true", help="処方明細のブック・CSV も書く")
    ap.add_argument("--export", help="ツールが読む形式で書き出す月（YYYY-MM）。all＝期間の全月。省略時は書かない")
    ap.add_argument("--export-days", type=int, default=0, help="その月の初めから何日分だけ書くか（0＝月まるごと）")
    ap.add_argument("--forecast-days", type=int, default=0, help="期間の翌日からの予測データ（予約から立つ予定）を何日分書くか")
    a = ap.parse_args()
    start, end = dt.date.fromisoformat(a.start), dt.date.fromisoformat(a.end)
    visits, plan = simulate(start, end + dt.timedelta(days=a.forecast_days), a.per_day)
    visits = [(d, s_) for d, s_ in visits if d <= end]
    if a.rx or not a.export:
        folder, by_day, agg = write_period(start, end, visits, plan)
        summary(start, end, visits, by_day, agg)
        print("出力：", folder)
    if a.export:
        mov, stock_eod, pending, avg, last_move = ledger(visits, start, end)
        if a.export == "all":
            months = []
            m = dt.date(start.year, start.month, 1)
            while m <= end:
                months.append(m); m = dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1)
        else:
            y, mo = map(int, a.export.split("-")); months = [dt.date(y, mo, 1)]
        for m in months:
            es, ee = m, dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1) - dt.timedelta(days=1)
            if a.export_days:
                ee = min(ee, es + dt.timedelta(days=a.export_days - 1))
            es, ee = max(es, start), min(ee, end)
            pp, sp, n_sheets, n_rows = write_exports(es, ee, plan, mov, stock_eod, pending)
            print(f"書き出し：{es:%Y/%m/%d}～{ee:%Y/%m/%d}  予定 {os.path.basename(pp)}（日別 {n_sheets} 枚）  出納 {os.path.basename(sp)}（{n_rows} 行）")
        for lp in write_system_lists(end, stock_eod, mov, last_move):
            print(f"  システムリスト：{lp}")
        ip = write_init_tables(start, end, end + dt.timedelta(days=1))
        print(f"  手入力テーブル：{ip}")
        if a.forecast_days:
            fs, fe = end + dt.timedelta(days=1), end + dt.timedelta(days=a.forecast_days)
            fp = write_forecast(fs, fe, end, plan, stock_eod, pending)
            print(f"  未来の予測：{fp}（{fs:%Y/%m/%d}～{fe:%Y/%m/%d}）")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
