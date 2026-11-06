# 組み立てた txt（dist/PowerPivot_ABC_<版>.txt）のクエリを、Excel を外から操作して分析用ブックに入れる
#   ・クエリ ⓪〜㉑ を書き込む（同名があれば置き換え、なければ追加）
#   ・T_／M_ の表をデータモデルへ、R_（A〜E）をシートの表へ、⑳㉑ を接続のみで作る
#   ・すべて更新し、リレーションシップ 10 本を作って保存する
# 使い方（このフォルダの親＝プロジェクトのフォルダで）：
#   pwsh -File tools\excel_install.ps1 -Book <ブック.xlsx> [-Txt dist\PowerPivot_ABC_0.3.0.txt]
#        [-Folder <実績フォルダ> -SuiFolder <出納フォルダ> -OrderDate 2026/10/01]
#   -Folder などを渡すと設定テーブルの値も書き換える（渡さなければ今の値のまま）
# 注意：
#   ・開発用（自宅 PC）。店の PC では使わない（マクロ・外部プログラムを使わない運用のため）
#   ・ブックには手入力テーブル（設定ほか）が先にあること（dist/PowerPivot_初期テーブル作成_<版>.txt の手順）
#   ・クエリを書き換えたあとは、Excel で「クエリのオプション → 現在のブック → プライバシー → 無視」を
#     選び直して保存する。Power Query はこの許可をクエリの内容と結び付けて保存するため、
#     書き換え後に開き直すと許可が無効になり、Formula.Firewall のエラーになる（0.3.0 で確認）
#   ・データモデルへの読み込みは Connections.Add2 の「コマンドの種類 6・コマンド "クエリ名"」で行う
#     （SQL 形式だと Excel が応答しなくなる。0.3.0 で確認）
param(
    [Parameter(Mandatory = $true)][string]$Book,
    [string]$Txt,
    [string]$Folder,
    [string]$SuiFolder,
    [string]$OrderDate,
    [string]$KigenFolder,       # 期限切迫フォルダパス（期限切れリスト）
    [string]$KeikaFolder,       # 経過措置フォルダパス（経過措置薬品リスト）
    [string]$FudoFolder,        # 不動品フォルダパス（不動品リスト）
    [switch]$SettingsOnly,      # 設定テーブルだけ書き換えて保存する（クエリ・接続は触らない）
    [string[]]$OnlyQueries      # 指定したクエリだけ書き換え、そのシートの表を更新して保存する（モデル・リレーションシップは触らない）
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ROOT = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $Txt) {
    $Txt = Get-ChildItem (Join-Path $ROOT "dist") -Filter "PowerPivot_ABC_*.txt" | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty FullName
}
Write-Host ("txt: " + $Txt)

# ---- txt をクエリごとに分ける（「// ◯ クエリ名：…」の見出しの 1 行上「// ====」から次の見出しの手前まで） ----
$lines = (Get-Content -Raw -Encoding UTF8 $Txt) -replace "`r`n", "`n" -split "`n"
$heads = @(); for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match '^// \S クエリ名：') { $heads += $i } }
$queries = @()
for ($k = 0; $k -lt $heads.Count; $k++) {
    $i = $heads[$k]
    $name = ([regex]::Match($lines[$i], '^// \S クエリ名：([^（]+)')).Groups[1].Value.Trim()
    $end = if ($k + 1 -lt $heads.Count) { $heads[$k + 1] - 1 } else { $lines.Count }
    $code = (($lines[($i - 1)..($end - 1)]) -join "`n").TrimEnd("`n")
    $queries += [pscustomobject]@{ name = $name; code = $code }
}
Write-Host ("クエリ " + $queries.Count + " 本")

$model  = @("T_日次", "T_月次", "M_日付", "M_薬品", "M_理由コード", "T_判定理由", "T_ABC月次", "T_まとまり月次", "T_ロット")
$connOnly = @("経過措置該当", "不動品リスト")
$sheets = [ordered]@{ "R_発注点見直し" = "A_発注点"; "R_不動在庫" = "B_不動在庫"; "R_経過措置" = "C_経過措置"; "R_出荷調整管理票" = "D_出荷調整"; "R_極小需要" = "E_極小需要"; "読込の確認" = "読込の確認" }
$rels = @(
    @("T_日次", "日付", "M_日付", "日付"), @("T_月次", "月", "M_日付", "日付"),
    @("T_日次", "薬品名", "M_薬品", "薬品名"), @("T_月次", "薬品名", "M_薬品", "薬品名"),
    @("T_判定理由", "薬品名", "M_薬品", "薬品名"), @("T_判定理由", "コード", "M_理由コード", "コード"),
    @("T_ABC月次", "月", "M_日付", "日付"), @("T_まとまり月次", "月", "M_日付", "日付"),
    @("T_まとまり月次", "薬品名", "M_薬品", "薬品名"), @("T_ロット", "薬品名", "M_薬品", "薬品名"))
function ConnStr($q) { "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$q;Extended Properties=`"`"" }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open((Resolve-Path $Book))     # Excel は相対パスを自分の作業フォルダで解決するので絶対パスにする

    # ---- 設定（渡された値だけ書き換える） ----
    if ($Folder -or $SuiFolder -or $OrderDate -or $KigenFolder -or $KeikaFolder -or $FudoFolder) {
        $lo = $wb.Worksheets.Item("設定").ListObjects.Item("設定")
        for ($i = 1; $i -le $lo.ListRows.Count; $i++) {
            $r = $lo.ListRows.Item($i).Range; $k = [string]$r.Cells.Item(1, 1).Value2; $c = $r.Cells.Item(1, 2)
            if ($k -eq "フォルダパス" -and $Folder) { $c.Value2 = $Folder }
            if ($k -eq "出納フォルダパス" -and $SuiFolder) { $c.Value2 = $SuiFolder }
            if ($k -eq "発注日" -and $OrderDate) { $c.Formula = ""; try { $c.NumberFormatLocal = "yyyy/mm/dd" } catch {}; $c.Value2 = [DateTime]$OrderDate }
            if ($k -eq "期限切迫フォルダパス" -and $KigenFolder) { $c.Value2 = $KigenFolder }
            if ($k -eq "経過措置フォルダパス" -and $KeikaFolder) { $c.Value2 = $KeikaFolder }
            if ($k -eq "不動品フォルダパス" -and $FudoFolder) { $c.Value2 = $FudoFolder }
        }
        Write-Host "設定を書き換えた"
    }
    if ($SettingsOnly) { $wb.Save(); $wb.Close($false); Write-Host "保存しました（-SettingsOnly）"; return }

    # ---- クエリ ----
    $have = @{}; foreach ($q in $wb.Queries) { $have[$q.Name] = $true }
    if ($OnlyQueries) {
        $OnlyQueries = @($OnlyQueries | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ })   # -File 経由では「a,b,c」が 1 つの文字列で来る
        Write-Host ("対象: " + ($OnlyQueries -join " / "))
        foreach ($q in $queries) {
            if ($OnlyQueries -notcontains $q.name) { continue }
            $wb.Queries.Item($q.name).Formula = $q.code; Write-Host ("更新 " + $q.name)
            if ($sheets.Contains($q.name)) {
                $lo = $wb.Worksheets.Item($sheets[$q.name]).ListObjects.Item(1); $lo.QueryTable.BackgroundQuery = $false
                try { $lo.QueryTable.Refresh($false) | Out-Null; Write-Host ("  " + $sheets[$q.name] + "：" + $lo.ListRows.Count + " 行") } catch { Write-Host ("※ 失敗 " + $q.name + "：" + $_.Exception.Message) }
            }
        }
        $wb.Save(); $wb.Close($false); Write-Host "保存しました（-OnlyQueries）"; return
    }
    foreach ($q in $queries) {
        if ($have.ContainsKey($q.name)) { $wb.Queries.Item($q.name).Formula = $q.code; Write-Host ("更新 " + $q.name) }
        else { $wb.Queries.Add($q.name, $q.code) | Out-Null; Write-Host ("追加 " + $q.name) }
    }

    # ---- 接続 ----
    $conn = @{}; foreach ($c in $wb.Connections) { $conn[$c.Name] = $c }
    $inModel = @{}; foreach ($mt in $wb.Model.ModelTables) { $inModel[$mt.Name] = $true }
    foreach ($q in $model) {
        if ($inModel.ContainsKey($q)) { continue }
        $cn = "クエリ - " + $q
        if ($conn.ContainsKey($cn)) { $conn[$cn].Delete() }
        Write-Host ("モデルに読み込み中… " + $q)
        $wb.Connections.Add2($cn, "ブック内の '" + $q + "' クエリへの接続です。", (ConnStr $q), ('"' + $q + '"'), 6, $true, $false) | Out-Null
        Write-Host ("モデルに読み込み " + $q)
    }
    foreach ($q in $connOnly) {
        $cn = "クエリ - " + $q
        if (-not $conn.ContainsKey($cn)) {
            $wb.Connections.Add2($cn, "ブック内の '" + $q + "' クエリへの接続です。", (ConnStr $q), "SELECT * FROM [" + $q + "]", 2, $false, $false) | Out-Null
            Write-Host ("接続のみ " + $q)
        }
    }
    $exist = @{}; foreach ($s in $wb.Worksheets) { $exist[$s.Name] = $true }
    $last = $wb.Worksheets.Item($wb.Worksheets.Count)
    foreach ($q in $sheets.Keys) {
        $sn = $sheets[$q]
        if ($exist.ContainsKey($sn)) { continue }
        $s = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $last); $s.Name = $sn; $last = $s
        $lo2 = $s.ListObjects.Add(0, (ConnStr $q), [System.Reflection.Missing]::Value, 1, $s.Range("A1"))
        $qt = $lo2.QueryTable
        $qt.CommandType = 2; $qt.CommandText = "SELECT * FROM [" + $q + "]"; $qt.BackgroundQuery = $false; $qt.AdjustColumnWidth = $true
        $lo2.Name = $q
        try { $qt.WorkbookConnection.Name = "クエリ - " + $q } catch {}
        try { $qt.Refresh($false) | Out-Null; Write-Host ("シートに読み込み " + $q + " → " + $sn + "（" + $lo2.ListRows.Count + " 行）") }
        catch { Write-Host ("※ 失敗 " + $q + " → " + $sn + "：" + $_.Exception.Message) }
    }

    # ---- すべて更新 ----
    foreach ($c in $wb.Connections) { try { if ($c.Type -eq 1) { $c.OLEDBConnection.BackgroundQuery = $false } } catch {} }
    try { $wb.RefreshAll(); $xl.CalculateUntilAsyncQueriesDone(); Write-Host "すべて更新：完了" }
    catch { Write-Host ("※ すべて更新で失敗：" + $_.Exception.Message) }

    # ---- リレーションシップ ----
    $m = $wb.Model
    $existing = @(); foreach ($r in $m.ModelRelationships) { $existing += ($r.ForeignKeyTable.Name + "." + $r.ForeignKeyColumn.Name + ">" + $r.PrimaryKeyTable.Name + "." + $r.PrimaryKeyColumn.Name) }
    foreach ($r in $rels) {
        $key = $r[0] + "." + $r[1] + ">" + $r[2] + "." + $r[3]
        if ($existing -contains $key) { continue }
        try {
            $fk = $m.ModelTables.Item($r[0]).ModelTableColumns.Item($r[1]); $pk = $m.ModelTables.Item($r[2]).ModelTableColumns.Item($r[3])
            $m.ModelRelationships.Add($fk, $pk) | Out-Null; Write-Host ("リレーションシップ " + $key)
        } catch { Write-Host ("※ リレーションシップ失敗 " + $key + "：" + $_.Exception.Message) }
    }

    foreach ($mt in $m.ModelTables) { Write-Host ("  モデル " + $mt.Name + "：" + $mt.RecordCount + " 行") }
    foreach ($q in $sheets.Keys) { try { Write-Host ("  表 " + $q + "：" + $wb.Worksheets.Item($sheets[$q]).ListObjects.Item(1).ListRows.Count + " 行") } catch { Write-Host ("  表 " + $q + "：なし") } }
    $wb.Worksheets.Item("設定").Activate() | Out-Null
    $wb.Save(); $wb.Close($false)
    Write-Host "保存しました（このあと Excel でプライバシーの「無視」を選び直して保存すること）"
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
