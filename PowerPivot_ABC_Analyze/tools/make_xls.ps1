# 検証用：架空データの .xlsx を、Excel を外から操作して本物の .xls（Excel 97-2003 形式）に変換する
#   エリシアSの実物は .xls なので、Excel での確認は .xls でも必ず行う（.xlsx だけでは旧形式の癖が出ない）
# 使い方（プロジェクトのフォルダで）：
#   pwsh -File tools\make_xls.ps1 [-Src tests\out] [-Dst tests\out\形式別\xls]
#   Src の下の 実績・出納・システムリスト にある .xlsx を、同じフォルダ構成で Dst に .xls として保存する
# 注意：開発用（自宅 PC）。店の PC では使わない
param(
    [string]$Src,
    [string]$Dst,
    [string[]]$Folders = @("実績", "出納", "システムリスト")
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ROOT = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $Src) { $Src = Join-Path $ROOT "tests\out" }
if (-not $Dst) { $Dst = Join-Path $Src "形式別\xls" }
$Src = (Resolve-Path $Src).Path
New-Item -ItemType Directory -Force $Dst | Out-Null
$Dst = (Resolve-Path $Dst).Path
$XLS = 56          # xlExcel8（Excel 97-2003 ブック）
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$n = 0
try {
    foreach ($sub in $Folders) {
        $from = Join-Path $Src $sub
        if (-not (Test-Path $from)) { continue }
        foreach ($f in Get-ChildItem $from -Recurse -Filter "*.xlsx") {
            if ($f.Name.StartsWith("~$")) { continue }
            $rel = $f.FullName.Substring($Src.Length).TrimStart("\")
            $out = Join-Path $Dst ([IO.Path]::ChangeExtension($rel, ".xls"))
            New-Item -ItemType Directory -Force (Split-Path -Parent $out) | Out-Null
            if (Test-Path $out) { Remove-Item $out -Force }
            $wb = $xl.Workbooks.Open($f.FullName, 0, $true)
            try { $wb.CheckCompatibility = $false } catch { Write-Host ("  互換性チェックを止められない：" + $_.Exception.Message) }
            $wb.SaveAs($out, $XLS)
            $wb.Close($false)
            $n++
        }
    }
    Write-Host ("変換：" + $n + " ファイル → " + $Dst)
}
catch { Write-Host ("※ 失敗：" + $_.Exception.Message); throw }
finally { $xl.Quit() }
