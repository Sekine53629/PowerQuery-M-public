# 検証用：ブックのクエリの結果を CSV に書き出す（ブックは保存しない）
#   データモデルの中身は外から読めないので、同じクエリを一時的にシートへ読み込み、値を CSV にして閉じる
# 使い方（プロジェクトのフォルダで）：
#   pwsh -File tools\excel_dump.ps1 -Book prototype\ABC分析.xlsx [-Out tests\out_excel] [-Queries M_薬品,T_月次,T_判定理由,T_ABC月次]
#   python tests\compare_excel.py tests\out_excel      ← 期待値（tests/out/期待値.xlsx）と突き合わせ
# 注意：
#   ・本体のブックをそのまま開く（保存しないので中身は変わらない）。控えにコピーすると、Excel が
#     「信頼済みドキュメント」として覚えているのは元のパスだけなので、プライバシーの「無視」が効かず
#     Formula.Firewall のエラーになる（0.3.0 で確認）
#   ・CSV は UTF-8（BOM 付き）。日付は yyyy-MM-dd、真偽は True/False、数値はそのまま
param(
    [Parameter(Mandatory = $true)][string]$Book,
    [string]$Out,
    [string[]]$Queries = @("M_薬品", "T_月次", "T_判定理由", "T_ABC月次"),
    [switch]$Touch        # 読み込む前にクエリ 1 本を同じ内容で書き直す（同じセッションで書き換えた直後だけ Formula.Firewall が出ない現象への回避。保存しないので中身は変わらない）
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ROOT = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $Out) { $Out = Join-Path $ROOT "tests\out_excel" }
New-Item -ItemType Directory -Force $Out | Out-Null
function ConnStr($q) { "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$q;Extended Properties=`"`"" }
function CsvCell($v) {
    if ($null -eq $v) { return "" }
    if ($v -is [DateTime]) { return $v.ToString("yyyy-MM-dd") }
    if ($v -is [bool]) { return $(if ($v) { "True" } else { "False" }) }
    if ($v -is [double] -or $v -is [int]) { return ([double]$v).ToString("R", [Globalization.CultureInfo]::InvariantCulture) }
    $s = [string]$v
    if ($s -match '[",\r\n]') { return '"' + $s.Replace('"', '""') + '"' }
    return $s
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Book))
    if ($Touch) { $q0 = $wb.Queries.Item("M_理由コード"); $q0.Formula = $q0.Formula; Write-Host "クエリを同じ内容で書き直した（-Touch）" }
    $last = $wb.Worksheets.Item($wb.Worksheets.Count)
    foreach ($q in $Queries) {
        $s = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $last); $s.Name = "chk_" + $q; $last = $s
        $lo = $s.ListObjects.Add(0, (ConnStr $q), [System.Reflection.Missing]::Value, 1, $s.Range("A1"))
        $qt = $lo.QueryTable; $qt.CommandType = 2; $qt.CommandText = "SELECT * FROM [" + $q + "]"; $qt.BackgroundQuery = $false
        try { $qt.Refresh($false) | Out-Null }
        catch { Write-Host ("※ 失敗 " + $q + "：" + $_.Exception.Message); continue }
        $vals = $lo.Range.Value()        # 見出し行を含む 2 次元配列（日付は DateTime で返る）
        $rows = $vals.GetLength(0); $cols = $vals.GetLength(1)
        $sb = New-Object System.Text.StringBuilder
        for ($r = 1; $r -le $rows; $r++) {
            $line = @(); for ($c = 1; $c -le $cols; $c++) { $line += (CsvCell $vals[$r, $c]) }
            [void]$sb.AppendLine(($line -join ","))
        }
        $path = Join-Path $Out ($q + ".csv")
        [IO.File]::WriteAllText($path, $sb.ToString(), (New-Object System.Text.UTF8Encoding $true))
        Write-Host ("" + $q + "：" + ($rows - 1) + " 行 → " + $path)
    }
    $wb.Close($false)
    Write-Host "ブックは保存していない"
}
finally { $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
