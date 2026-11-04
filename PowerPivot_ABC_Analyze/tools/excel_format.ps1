# 月次レポートのシート A〜E の体裁を整える（Excel を外から操作。列幅・書式・条件付き書式・スパークライン・印刷設定）
#   ・体裁はテーブル（ListObject）の列に付けるので、［すべて更新］で行が増減しても残る
#   ・スパークライン（A：月01〜月12、D：構成比_月01〜12）だけは、更新で行が増えたときにその行には付かない。
#     そのときはこのスクリプトをもう一度実行する（何度実行しても同じ結果になる）
# 使い方（プロジェクトのフォルダで）：
#   pwsh -File tools\excel_format.ps1 -Book prototype\ABC分析.xlsx [-Pdf <出力.pdf>]
#   -Pdf を付けると A〜E を 1 本の PDF に印刷して確認できる
# 注意：開発用（自宅 PC）。ブックは保存する（クエリ・接続は触らない）
param(
    [Parameter(Mandatory = $true)][string]$Book,
    [string]$Pdf
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ---- 色（薄い塗り） ----
$RED    = 0xD9E9FD   # BGR：薄い赤（注意）
$BLUE   = 0xF7EBDD   # 薄い青（減少）
$YELLOW = 0xCCF2FF   # 薄い黄（経過措置）
$GRAY   = 0x808080   # 文字の灰色
$DARKRED = 0x0000C0

# ---- 列の幅と表示形式（見出し名で指定。ないものは無視） ----
$WIDTH = @{
    "需要のまとまり" = 26; "薬品名" = 30; "代表の薬" = 30; "略称" = 11; "棚番" = 7; "代替の屋号" = 26;
    "処置の案" = 48; "切替・転送の案" = 60; "理由" = 34; "状態" = 24; "持ち方" = 14; "役割" = 16; "需要の型" = 11;
    "区分" = 8; "前月の区分" = 6; "今月の区分" = 6; "区分変更" = 5; "屋号数" = 5; "行順" = 4; "疑い" = 8;
    "期間" = 15; "先発後発" = 7; "供給状況" = 8; "期限切迫" = 5; "発注除外" = 5; "注意" = 5
}
$FMT = @{
    "直近の払出" = "#,##0"; "前の払出" = "#,##0"; "発注点の目安" = "#,##0"; "発注量の目安" = "#,##0"; "在庫" = "#,##0";
    "余った量" = "#,##0"; "余る量" = "#,##0"; "期限までに使い切れる量" = "#,##0"; "先月の仕入" = "#,##0"; "まとまりの払出" = "#,##0";
    "年間の予定患者数" = "#,##0"; "在庫日数" = "#,##0"; "直近の在庫日数" = "#,##0"; "システム不動日数" = "#,##0"; "残り日数" = "#,##0";
    "動きのない月数" = "0"; "動いた月数" = "0"; "屋号数" = "0"; "行順" = "0";
    "直近の1日平均払出" = "0.0"; "1日平均払出" = "0.0"; "平均間隔" = "0.0";
    "変化" = "+0%;-0%;0%"; "構成比" = "0%";
    "最後に動いた月" = "yyyy/mm"; "システム最終払出日" = "yyyy/mm/dd"; "経過措置期限" = "yyyy/mm/dd"
}
$WRAP = @("処置の案", "切替・転送の案", "理由", "状態")

# 条件付き書式：シートごとに (列, 種類, 値, 塗り, 文字色, 太字)。種類：eq（等しい）/ contains（含む）/ ge（以上）/ le（以下）/ lt（未満）/ true
$CF = @{
    "A_発注点" = @(
        @("区分変更", "eq", "○", $YELLOW, $null, $true),
        @("変化", "ge", 0.2, $RED, $null, $false), @("変化", "le", -0.2, $BLUE, $null, $false),
        @("持ち方", "contains", "持たない", $null, $GRAY, $false))
    "B_不動在庫" = @(
        @("区分", "eq", "不動", $RED, $null, $false),
        @("処置の案", "contains", "返品", $null, $null, $true))
    "C_経過措置" = @(
        @("期限切迫", "eq", "○", $RED, $null, $false),
        @("残り日数", "lt", 0, $null, $DARKRED, $true),
        @("余る量", "ge", 1, $null, $null, $true))
    "D_出荷調整" = @(
        @("疑い", "eq", "経過措置", $YELLOW, $null, $false), @("疑い", "eq", "疑い", $RED, $null, $false),
        @("役割", "eq", "推奨ジェネリック", $null, $null, $true),
        @("切替・転送の案", "contains", "【警告】", $null, $DARKRED, $true))
    "E_極小需要" = @(
        @("注意", "eq", "○", $RED, $null, $false),
        @("発注除外", "eq", "○", $null, $GRAY, $false))
}
# スパークライン：シート → (元の列の先頭, 末尾, 新しい列の名前, 種類 1=折れ線 2=縦棒)
$SPARK = @{
    "A_発注点"   = @("月01", "月12", "12ヶ月の払出", 2)
    "D_出荷調整" = @("構成比_月01", "構成比_月12", "12ヶ月の構成比", 1)
}
$TITLE = @{ "A_発注点" = "A 発注点の切り替え提案"; "B_不動在庫" = "B 不動在庫と転送・返品推奨"; "C_経過措置" = "C 経過措置・期限切迫品";
            "D_出荷調整" = "D 出荷調整の管理票"; "E_極小需要" = "E 要管理の極小需要薬" }

function ColIndex($lo, $name) { for ($i = 1; $i -le $lo.ListColumns.Count; $i++) { if ($lo.ListColumns.Item($i).Name -eq $name) { return $i } }; return 0 }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Book))
    foreach ($sn in $TITLE.Keys) {
        $ws = $null; try { $ws = $wb.Worksheets.Item($sn) } catch { Write-Host ("シートなし " + $sn); continue }
        $lo = $ws.ListObjects.Item(1)
        $ws.Cells.Font.Size = 9
        $lo.TableStyle = "TableStyleLight1"
        $lo.ShowTableStyleRowStripes = $true
        $hdr = $lo.HeaderRowRange; $hdr.WrapText = $true; $hdr.VerticalAlignment = -4108; $hdr.Font.Bold = $true

        # 列幅・表示形式・折り返し
        for ($i = 1; $i -le $lo.ListColumns.Count; $i++) {
            $c = $lo.ListColumns.Item($i); $n = $c.Name; $rng = $c.Range
            if ($WIDTH.ContainsKey($n)) { $rng.EntireColumn.ColumnWidth = $WIDTH[$n] } elseif ($n -match '^(構成比_)?月\d\d$') { $rng.EntireColumn.ColumnWidth = 5 } else { $rng.EntireColumn.ColumnWidth = 9 }
            if ($FMT.ContainsKey($n)) { $rng.NumberFormatLocal = $FMT[$n]; $rng.HorizontalAlignment = -4152 }
            if ($WRAP -contains $n) { $rng.WrapText = $true; $rng.VerticalAlignment = -4160 }
        }
        $lo.Range.VerticalAlignment = -4160

        # スパークライン（元の 12 列は隠す）
        if ($SPARK.ContainsKey($sn)) {
            $s0, $s1, $newName, $kind = $SPARK[$sn]
            $i0 = ColIndex $lo $s0; $i1 = ColIndex $lo $s1
            if ($i0 -gt 0 -and $i1 -gt 0) {
                $iNew = ColIndex $lo $newName
                if ($iNew -eq 0) { $col = $lo.ListColumns.Add(); $col.Name = $newName; $iNew = $col.Index }
                $col = $lo.ListColumns.Item($iNew); $col.Range.EntireColumn.ColumnWidth = 18
                $body = $col.DataBodyRange
                if ($null -ne $body) {
                    $body.SparklineGroups.Clear()
                    $src = $ws.Range($lo.ListColumns.Item($i0).DataBodyRange, $lo.ListColumns.Item($i1).DataBodyRange)
                    $g = $body.SparklineGroups.Add($kind, $src.Address($false, $false, 1, $true))
                    $g.DisplayHidden = $true
                    $g.SeriesColor.Color = 0xA56E3B      # 青系
                    if ($kind -eq 2) { $g.Points.Lastpoint.Visible = $true; $g.Points.Lastpoint.Color.Color = 0x3B6EE7 }
                }
                for ($i = $i0; $i -le $i1; $i++) { $lo.ListColumns.Item($i).Range.EntireColumn.Hidden = $true }
                Write-Host ("  " + $sn + "：スパークライン " + $newName)
            }
        }

        # 条件付き書式（テーブルの列全体に付けるので、行が増えても効く）
        if ($CF.ContainsKey($sn)) {
            foreach ($rule in $CF[$sn]) {
                $n, $kind, $val, $fill, $font, $bold = $rule
                $i = ColIndex $lo $n; if ($i -eq 0) { continue }
                $rng = $lo.ListColumns.Item($i).DataBodyRange; if ($null -eq $rng) { continue }
                $rng.FormatConditions.Delete()
                switch ($kind) {
                    "eq"       { $fc = $rng.FormatConditions.Add(1, 3, ('="' + $val + '"')) }      # xlCellValue, xlEqual
                    "ge"       { $fc = $rng.FormatConditions.Add(1, 7, [string]$val) }             # xlGreaterEqual
                    "le"       { $fc = $rng.FormatConditions.Add(1, 8, [string]$val) }             # xlLessEqual
                    "lt"       { $fc = $rng.FormatConditions.Add(1, 6, [string]$val) }             # xlLess
                    "contains" { $m = [System.Reflection.Missing]::Value; $fc = $rng.FormatConditions.Add(9, $m, $m, $m, $val, 0) }   # xlTextString, String, xlContains
                    "true"     { $fc = $rng.FormatConditions.Add(1, 3, "=TRUE") }
                }
                if ($fill) { $fc.Interior.Color = $fill }
                if ($font) { $fc.Font.Color = $font }
                if ($bold) { $fc.Font.Bold = $true }
            }
        }

        # 印刷：A4 横、横 1 ページ、見出し行を各ページに、ヘッダーにタイトルと日付
        $ws.Activate() | Out-Null
        $ws.Range("A2").Select() | Out-Null
        $xl.ActiveWindow.FreezePanes = $false; $xl.ActiveWindow.FreezePanes = $true
        $xl.PrintCommunication = $false
        $ps = $ws.PageSetup
        $ps.PaperSize = 9; $ps.Orientation = 2; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false
        $ps.PrintTitleRows = '$1:$1'; $ps.CenterHorizontally = $true
        $ps.LeftMargin = 28; $ps.RightMargin = 28; $ps.TopMargin = 42; $ps.BottomMargin = 36
        $ps.LeftHeader = "&""游ゴシック,太字""&11" + $TITLE[$sn]; $ps.RightHeader = "作成 &D"; $ps.CenterFooter = "&P / &N"
        $xl.PrintCommunication = $true
        Write-Host ($sn + "：体裁を付けた（" + $lo.ListRows.Count + " 行）")
    }
    if ($Pdf) {
        $wb.Worksheets.Item(@("A_発注点", "B_不動在庫", "C_経過措置", "D_出荷調整", "E_極小需要")).Select() | Out-Null
        $xl.ActiveSheet.ExportAsFixedFormat(0, [IO.Path]::GetFullPath($Pdf)) | Out-Null
        $wb.Worksheets.Item("A_発注点").Select() | Out-Null
        Write-Host ("PDF: " + $Pdf)
    }
    $wb.Worksheets.Item("A_発注点").Activate() | Out-Null
    $wb.Save(); $wb.Close($false); Write-Host "保存しました"
}
finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
