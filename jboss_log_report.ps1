<#
.SYNOPSIS
    JBoss EAP ログ検索結果 Excel レポート作成  (Windows 11 / PowerShell 5.1 以降)

.DESCRIPTION
    jboss_log_search.sh / jboss_access_log_search.sh / jboss_gc_log_search.sh の
    実行結果を記録した TeraTerm ログを読み込み、Excel (.xlsx) レポートを出力する (Excel 本体は不要)
      出力       : サマリー シート (検索一覧 / 検索条件 / 集計) + 検索ごとの明細シート (1 件 1 行)
      文字コード : 自動判定 (BOM 付き UTF-8 / UTF-16, UTF-8, CP932)。-Encoding で明示指定も可
      TeraTerm   : エスケープシーケンス (色) / 行頭タイムスタンプ / プロンプト等は自動で除去
    ※ 本ファイルは BOM 付き UTF-8 で保存すること (PowerShell 5.1 で日本語を正しく読み込むため)

.PARAMETER Path
    TeraTerm ログ (複数指定 / ワイルドカード可)

.PARAMETER OutFile
    出力ファイル。省略時: 入力 1 件 = 入力と同じ場所に <入力名>.xlsx
                          入力 複数 = カレントに jboss_log_report_YYYYMMDD_HHMMSS.xlsx

.PARAMETER Encoding
    入力の文字コード (Auto / UTF8 / UTF16LE / UTF16BE / CP932)。既定は Auto (自動判定)

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\jboss_log_report.ps1 C:\TeraTerm\log\teraterm.log

.EXAMPLE
    .\jboss_log_report.ps1 .\log\*.log -OutFile .\report.xlsx
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]] $Path,
    [string] $OutFile,
    [ValidateSet('Auto', 'UTF8', 'UTF16LE', 'UTF16BE', 'CP932')]
    [string] $Encoding = 'Auto'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Fail([string]$msg, [int]$code = 1) { [Console]::Error.WriteLine("[ERROR] $msg"); exit $code }

# --- 入力ファイル / 出力ファイル ---
$inputs = New-Object System.Collections.Generic.List[string]
foreach ($p in $Path) {
    $items = @(Get-Item -Path $p -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer })
    if ($items.Count -eq 0) { $items = @(Get-Item -LiteralPath $p -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer }) }
    if ($items.Count -eq 0) { Fail "ファイルを読み込めません : $p" }
    foreach ($i in $items) { $inputs.Add($i.FullName) }
}
if (-not $OutFile) {
    if ($inputs.Count -eq 1) { $OutFile = [IO.Path]::ChangeExtension($inputs[0], '.xlsx') }
    else { $OutFile = 'jboss_log_report_{0:yyyyMMdd_HHmmss}.xlsx' -f (Get-Date) }
}
$OutFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)

if ($PSVersionTable.PSEdition -eq 'Core') { [Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance) }

# --- 入力を文字列に変換 (BOM で UTF-8 / UTF-16 を判定、BOM なしは UTF-8 として不正なら CP932) ---
function Read-LogText([string]$file) {
    $b = [IO.File]::ReadAllBytes($file)
    switch ($Encoding) {
        'UTF8'    { $enc = New-Object Text.UTF8Encoding($false) }
        'UTF16LE' { $enc = [Text.Encoding]::Unicode }
        'UTF16BE' { $enc = [Text.Encoding]::BigEndianUnicode }
        'CP932'   { $enc = [Text.Encoding]::GetEncoding(932) }
        default {
            if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) { return [Text.Encoding]::UTF8.GetString($b, 3, $b.Length - 3) }
            if ($b.Length -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($b, 2, $b.Length - 2) }
            if ($b.Length -ge 2 -and $b[0] -eq 0xFE -and $b[1] -eq 0xFF) { return [Text.Encoding]::BigEndianUnicode.GetString($b, 2, $b.Length - 2) }
            try { return (New-Object Text.UTF8Encoding($false, $true)).GetString($b) } catch { $enc = [Text.Encoding]::GetEncoding(932) }
        }
    }
    $s = $enc.GetString($b)
    if ($s.Length -gt 0 -and $s[0] -eq [char]0xFEFF) { $s = $s.Substring(1) }
    $s
}

# --- 定義 ---
$INV    = [Globalization.CultureInfo]::InvariantCulture
$WS     = " `t".ToCharArray()
$MAXROW = 1000000
# 出力行の項目数 (最後の項目はメッセージ全体) / 続き行 (スタックトレース等) を格納する列
$NC = @{ 'server.log' = 6; 'access_log' = 7; 'gc.log' = 7 }
$DC = @{ 'server.log' = 10; 'gc.log' = 11 }
# 明細シートの列 : 見出し : 種別 : 最小幅 : 最大幅
$COLS = @{
    'server.log' = 'No:no:7:7|TIMESTAMP (JST):dtms:25:25|CTR:ctr:7:10|SERVICE:txt:11:20|TASK:ctr:11:14|LEVEL:lvl:8:8|CATEGORY:txt:14:45|THREAD:txt:14:30|MESSAGE:txt:40:100|DETAIL (続き行):txt:20:80'
    'access_log' = 'No:no:7:7|TIMESTAMP (JST):dts:21:21|CTR:ctr:7:10|SERVICE:txt:11:20|TASK:ctr:11:14|LEVEL:lvl:8:8|STS:code:7:7|CLIENT:txt:12:20|METHOD:ctr:10:10|URL:txt:30:70|BYTES:int:10:10|ACCESS LOG:txt:40:100'
    'gc.log'     = 'No:no:7:7|TIMESTAMP (JST):dtms:25:25|CTR:ctr:7:10|SERVICE:txt:11:20|TASK:ctr:11:14|LEVEL:lvl:8:8|PAUSE(ms):dec:12:12|TAGS:txt:10:22|GC ID:code:8:8|GC LOG:txt:40:100|DETAIL (続き行):txt:20:80'
}
# cellXfs の番号 (styles.xml と対応 / データ用は +1 で縞模様)
$S_TITLE = 1; $S_META = 2; $S_SEC = 3; $S_CAP = 4; $S_HEAD = 5; $S_LABEL = 6; $S_VALUE = 7
$S_TEXT = 8; $S_CTR = 10; $S_DTMS = 12; $S_DTS = 14; $S_INT = 16; $S_DEC = 18; $S_CODE = 20; $S_LINK = 22; $S_DIM = 24
$S_ERR = 26; $S_WRN = 27; $S_TTEXT = 28; $S_TINT = 29; $S_TDEC = 30; $S_INTB = 31

$reEsc   = [regex]'\x1B\][^\x07\x1B]*(?:\x07|\x1B\\)|\x1B\[[0-9;?<=>]*[ -/]*[@-~]|\x1B[()*+][0-9A-Za-z]|\x1B[ -/]*[0-~]|[\x00-\x08\x0B-\x1F\x7F]'
$reTs    = [regex]'^\[[^\]]*[0-9]:[0-9]{2}:[0-9]{2}[^\]]*\] ?'
$reTitle = [regex]'^  JBoss EAP (server\.log|access_log|gc\.log) 検索$'
$reHr    = [regex]'^=+$'
$reLr    = [regex]'^-+$'
$reData  = [regex]'^[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}(\.[0-9]{3})? \| '
$reCont  = [regex]'^      +\|( |$)'
$reCond  = [regex]'^    [^ ]'
$reCondC = [regex]'^      +[^ ]'
$reSrv   = [regex]'^\[([^\]]*)\] \(([^)]*)\) ?(.*)$'
$reCli   = [regex]'^[^ "]+'
$reReq   = [regex]'"([^"]*)"( +([0-9]{3}|-) +([^ ]+))?'
$reTag   = [regex]'^\[([a-z0-9,]+)\] (.*)$'
$reGcId  = [regex]'^GC\(([0-9]+)\) (.*)$'
$reLvErr = [regex]'^(ERROR|FATAL)$'
$reGrp   = [regex]'^  [^ ]+ / [^ ]+ '
$reSts   = [regex]'([0-9]{3}|-)=([0-9]+)'
$reTop   = [regex]'^ +[0-9]+\. '
$reFHdr  = [regex]'^ +HITS +ERROR +WARN .*FILE \(([0-9]+) 件\)'
$reFile  = [regex]'^ +([0-9]+) +([0-9]+) +([0-9]+) +(([0-9]+\.[0-9]+|-) +)?([^ ].*)$'
$reUnp   = [regex]' +時刻解析不可 ([0-9]+) 行'
$reNum   = [regex]'^-?[0-9]+(\.[0-9]+)?$'
$reDate  = [regex]'^([0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2})(\.[0-9]{3})?'
$reCtl   = [regex]'[\x00-\x08\x0B-\x1F\x7F]'
$reNarrow = [regex]'[ -~]'
$SEP     = [string[]]@(' | ')

# ---------------- 解析 ----------------
function Split-Sep([string]$s, [int]$n) {           # " | " 区切りで最大 n 項目に分割 (最後の項目は残り全部)
    $a = $s.Split($SEP, $n, [StringSplitOptions]::None)
    if ($a.Count -lt $n) { $a = [string[]]($a + @('') * ($n - $a.Count)) }
    , $a
}

function New-Block([string]$type, [string]$src) {
    @{
        Type = $type; Src = $src; Sheet = ''; Err = ''; SS = ''; FileCount = ''; Total = $null; EC = 0; WC = 0; Cut = $false
        Conds  = New-Object System.Collections.Generic.List[object]
        Notes  = New-Object System.Collections.Generic.List[string]
        Rows   = New-Object System.Collections.Generic.List[object]
        Groups = New-Object System.Collections.Generic.List[object]
        Status = New-Object System.Collections.Generic.List[object]
        Top    = New-Object System.Collections.Generic.List[object]
        Files  = New-Object System.Collections.Generic.List[object]
    }
}

function Add-Row($B, [string]$s) {
    $t = $B.Type; $n = $NC[$t]
    $F = $s.Split($SEP, $n, [StringSplitOptions]::None)  # Split-Sep と同じ (関数呼び出しを省略)
    if ($F.Count -lt $n) { $F = [string[]]($F + @('') * ($n - $F.Count)) }
    for ($i = 0; $i -lt $F.Count - 1; $i++) { $F[$i] = $F[$i].Trim($WS) }
    if ($B.Rows.Count -ge $MAXROW) {
        if (-not $B.Cut) { $B.Cut = $true; $B.Notes.Add("明細が $MAXROW 行を超えたため以降を省略しました") }
        return
    }
    $R = New-Object string[] 13
    $R[2] = $F[0]; $R[3] = $F[1]; $R[4] = $F[2]; $R[5] = $F[3]; $R[6] = $F[4]
    if ($t -eq 'server.log') {                       # [カテゴリ] (スレッド) メッセージ
        $m = $F[5]
        $x = $reSrv.Match($m)
        if ($x.Success) { $R[7] = $x.Groups[1].Value; $R[8] = $x.Groups[2].Value; $m = $x.Groups[3].Value }
        $R[9] = $m
    } elseif ($t -eq 'access_log') {                 # クライアント ... "メソッド URL プロトコル" ステータス バイト数
        $m = $F[6]; $R[7] = $F[5]
        $x = $reCli.Match($m)
        if ($x.Success) { $R[8] = $x.Value }
        $x = $reReq.Match($m)
        if ($x.Success) {
            $q = @($x.Groups[1].Value.Trim($WS) -split ' +')
            $R[9] = $q[0]; if ($q.Count -gt 1) { $R[10] = $q[1] }; $R[11] = $x.Groups[4].Value
        }
        $R[12] = $m
    } else {                                         # [タグ] GC(ID) メッセージ
        $m = $F[6]; $R[7] = $F[5]
        $x = $reTag.Match($m)
        if ($x.Success) { $R[8] = $x.Groups[1].Value; $m = $x.Groups[2].Value }
        $x = $reGcId.Match($m)
        if ($x.Success) { $R[9] = $x.Groups[1].Value; $m = $x.Groups[2].Value }
        $R[10] = $m
    }
    if ($reLvErr.IsMatch($F[4])) { $B.EC++ } elseif ($F[4] -ceq 'WARN') { $B.WC++ }
    $B.Rows.Add($R)
}

function Read-Summary($B, [string]$s) {
    if ($reLr.IsMatch($s)) { $B.SS = ''; return }
    if ($s.StartsWith('  集計 ', 'Ordinal')) { return }
    if ($s.StartsWith('  CONTAINER / SERVICE ', 'Ordinal')) { $B.SS = 'G'; return }
    if ($s.StartsWith('  停止時間 TOP', 'Ordinal')) { $B.SS = 'T'; return }
    if ($s.StartsWith('  TOTAL ', 'Ordinal')) { $B.Total = @(-split $s); return }
    $m = $reFHdr.Match($s)
    if ($m.Success) { $B.SS = 'F'; $B.FileCount = $m.Groups[1].Value; return }
    if ($s.Contains('時刻を解析できない行が')) { $B.Notes.Add($s.Trim($WS)); return }
    if ($s.StartsWith('  STATUS ', 'Ordinal')) { $B.SS = 'S'; $s = $s.Substring(9) }

    if ($B.SS -eq 'G' -and $reGrp.IsMatch($s)) {     # front / intra-web  HITS ERROR WARN [PAUSES FULL MAX TOTAL]
        $a = @(-split $s)
        $g = New-Object string[] ($a.Count - 1)
        $g[1] = $a[0] + ' / ' + $a[2]
        for ($i = 3; $i -lt $a.Count; $i++) { $g[$i - 1] = $a[$i] }
        $B.Groups.Add($g)
    } elseif ($B.SS -eq 'S') {                       # 200=10  404=2 ...
        foreach ($x in $reSts.Matches($s)) { $B.Status.Add(@($x.Groups[1].Value, $x.Groups[2].Value)) }
    } elseif ($B.SS -eq 'T' -and $reTop.IsMatch($s)) {   # 1. TIMESTAMP | CTR / SERVICE / TASK | PAUSE ms | GC LOG
        $P = Split-Sep $s 4
        $a = @($P[1].Split('/') + @('', '', ''))
        $B.Top.Add(@($reTop.Replace($P[0], '').Trim($WS), $a[0].Trim($WS), $a[1].Trim($WS), $a[2].Trim($WS),
                     ($P[2] -creplace 'ms *$', '').Trim($WS), $P[3]))
    } elseif ($B.SS -eq 'F') {                       # HITS ERROR WARN [MAX] FILE [時刻解析不可 N 行]
        $m = $reFile.Match($s)
        if ($m.Success) {
            $path = $m.Groups[6].Value; $unp = ''
            $u = $reUnp.Match($path)
            if ($u.Success) { $unp = $u.Groups[1].Value; $path = $path.Substring(0, $u.Index) }
            $B.Files.Add(@($m.Groups[1].Value, $m.Groups[2].Value, $m.Groups[3].Value, $m.Groups[5].Value, $unp, $path.Trim($WS)))
        }
    }
}

$blocks = New-Object System.Collections.Generic.List[object]
foreach ($file in $inputs) {
    $text = Read-LogText $file
    $st = 0; $B = $null
    foreach ($line in $text.Split("`n")) {
        $s = $reEsc.Replace($line, '')                # TeraTerm ログの制御文字・タイムスタンプを除去
        $s = $reTs.Replace($s, '').TrimEnd($WS)

        $m = $reTitle.Match($s)
        if ($m.Success) { $B = New-Block $m.Groups[1].Value $file; $blocks.Add($B); $st = 1; continue }
        if ($st -eq 0) { continue }
        if ($s.StartsWith('[SKIP] ', 'Ordinal')) { $B.Notes.Add('読込スキップ : ' + $s.Substring(7)); continue }
        $hr = $s.Length -ge 20 -and $reHr.IsMatch($s)
        $lr = $s.Length -ge 20 -and $reLr.IsMatch($s)

        if ($st -eq 1) {                              # 検索条件
            $p = $s.IndexOf(' : ', [StringComparison]::Ordinal)
            if ($hr) { $st = 0 }
            elseif ($lr) { $st = 2 }
            elseif ($s.Contains('FROM / TO の書式が不正です')) { $B.Err = $s.Trim($WS) }
            elseif ($reCond.IsMatch($s) -and $p -gt 0) { $B.Conds.Add(@($s.Substring(0, $p).Trim($WS), $s.Substring($p + 3).Trim($WS))) }
            elseif ($reCondC.IsMatch($s) -and $B.Conds.Count) { $c = $B.Conds[$B.Conds.Count - 1]; $c[1] = $c[1] + "`n" + $s.Trim($WS) }
            continue
        }
        if ($st -eq 2 -or $st -eq 3) {                # 列見出し
            if ($s.StartsWith('TIMESTAMP', 'Ordinal')) { $st = 3; continue }
            if ($lr -and $st -eq 3) { $st = 4; continue }
            if ($hr) { $st = 5; continue }
            if (-not $reData.IsMatch($s)) { continue }
            $st = 4
        }
        if ($st -eq 4) {                              # 明細
            if ($hr) { $st = 5 }
            elseif ($reData.IsMatch($s)) { Add-Row $B $s }
            elseif ($reCont.IsMatch($s) -and $B.Rows.Count -and $DC.ContainsKey($B.Type) -and -not $B.Cut) {
                $r = $B.Rows[$B.Rows.Count - 1]; $t = $DC[$B.Type]
                $i = $s.IndexOf('|') + 2
                $x = if ($i -le $s.Length) { $s.Substring($i) } else { '' }
                $r[$t] = if ($r[$t]) { $r[$t] + "`n" + $x } else { $x }
            }
            continue
        }
        if ($hr) { $st = 0; continue }                # 集計
        Read-Summary $B $s
    }
}
if ($blocks.Count -eq 0) { Fail 'JBoss EAP の検索結果が見つかりません (TeraTerm ログを確認してください)' 3 }

# ---------------- xlsx 出力 ----------------
$NS    = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$NSR   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$XMLH  = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>`n"
$NOW   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
$NOWZ  = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
$parts = [ordered]@{}

function ColName([int]$n) { $s = ''; while ($n -gt 0) { $s = [string][char](65 + ($n - 1) % 26) + $s; $n = [math]::Floor(($n - 1) / 26) }; $s }
$CLN = @('') + @(1..26 | ForEach-Object { ColName $_ })

function Esc([string]$s) { $reCtl.Replace($s, '').Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;') }
function Get-Width([string]$s) {                     # 表示幅の目安 (全角 = 2)
    $i = $s.IndexOf("`n"); if ($i -ge 0) { $s = $s.Substring(0, $i) }
    if ($s.Length -gt 200) { return 200 }
    $n = $s.Length
    foreach ($ch in $s.ToCharArray()) { if ([int]$ch -lt 32 -or [int]$ch -gt 126) { $n++ } }
    $n
}
function ConvertTo-Serial([string]$v) {              # "YYYY-MM-DD HH:MM:SS[.mmm]" -> Excel シリアル値
    $m = $reDate.Match($v)
    if (-not $m.Success) { return '' }
    $d = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($m.Groups[1].Value, 'yyyy-MM-dd HH:mm:ss', $INV, [Globalization.DateTimeStyles]::None, [ref]$d)) { return '' }
    if ($m.Groups[2].Success) { $d = $d.AddMilliseconds([int]$m.Groups[2].Value.Substring(1)) }
    $d.ToOADate().ToString('0.0000000000', $INV)
}
function Row-Open([int]$r, $ht) { $script:CR = $r; [void]$script:Buf.Append("<row r=""$r"" ht=""$ht"" customHeight=""1"">") }
function Row-Close { [void]$script:Buf.Append("</row>`n") }
function Cell-Str([int]$c, [string]$v, [int]$st) {   # 文字列セル
    $ref = $CLN[$c] + $script:CR
    if ($v -eq '') { [void]$script:Buf.Append("<c r=""$ref"" s=""$st""/>"); return }
    if ($v.Length -gt 32000) { $v = $v.Substring(0, 32000) + "`n...(32,000 文字を超えるため省略)" }
    [void]$script:Buf.Append("<c r=""$ref"" s=""$st"" t=""inlineStr""><is><t xml:space=""preserve"">").Append((Esc $v)).Append('</t></is></c>')
}
function Cell-Num([int]$c, [string]$v, [int]$st) {   # 数値セル (数値でなければ文字列)
    if (-not $reNum.IsMatch($v)) { Cell-Str $c $v $st; return }
    [void]$script:Buf.Append("<c r=""$($CLN[$c])$($script:CR)"" s=""$st""><v>$v</v></c>")
}
function Cell-Date([int]$c, [string]$v, [int]$st, [int]$alt) {
    $x = ConvertTo-Serial $v
    if ($x) { Cell-Num $c $x $st } else { Cell-Str $c $v $alt }
}
function Cell-Merge([int]$c1, [int]$c2, [string]$v, [int]$st) {   # 結合セル
    Cell-Str $c1 $v $st
    for ($c = $c1 + 1; $c -le $c2; $c++) { Cell-Str $c '' $st }
    if ($c2 -gt $c1) { $script:MG.Add("$($CLN[$c1])$($script:CR):$($CLN[$c2])$($script:CR)") }
}
function Cell-Link([int]$c, [string]$txt, [int]$st) {
    Cell-Str $c $txt $st
    $script:HL.Add(@("$($CLN[$c])$($script:CR)", "'$txt'!A1", $txt))
}
function Sheet-Open([int]$idx, [int]$frz, [int]$nc, $W) {
    $script:SheetIdx = $idx
    $script:Buf = New-Object System.Text.StringBuilder 65536
    $script:MG = New-Object System.Collections.Generic.List[string]
    $script:HL = New-Object System.Collections.Generic.List[object]
    [void]$script:Buf.Append("$XMLH<worksheet xmlns=""$NS"" xmlns:r=""$NSR""><sheetPr><pageSetUpPr fitToPage=""1""/></sheetPr>")
    [void]$script:Buf.Append("<sheetViews><sheetView workbookViewId=""0"" showGridLines=""0""$(if ($idx -eq 1) { ' tabSelected="1"' })>")
    if ($frz) { [void]$script:Buf.Append("<pane ySplit=""$frz"" topLeftCell=""A$($frz + 1)"" activePane=""bottomLeft"" state=""frozen""/><selection pane=""bottomLeft"" activeCell=""A$($frz + 1)"" sqref=""A$($frz + 1)""/>") }
    [void]$script:Buf.Append('</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18" customHeight="1"/><cols>')
    for ($i = 1; $i -le $nc; $i++) { [void]$script:Buf.Append("<col min=""$i"" max=""$i"" width=""$($W[$i])"" customWidth=""1""/>") }
    [void]$script:Buf.Append("</cols><sheetData>`n")
}
function Sheet-Close([string]$af) {
    [void]$script:Buf.Append('</sheetData>')
    if ($af) { [void]$script:Buf.Append("<autoFilter ref=""$af""/>") }
    if ($script:MG.Count) {
        [void]$script:Buf.Append("<mergeCells count=""$($script:MG.Count)"">")
        foreach ($m in $script:MG) { [void]$script:Buf.Append("<mergeCell ref=""$m""/>") }
        [void]$script:Buf.Append('</mergeCells>')
    }
    if ($script:HL.Count) {
        [void]$script:Buf.Append('<hyperlinks>')
        foreach ($h in $script:HL) { [void]$script:Buf.Append("<hyperlink ref=""$($h[0])"" location=""$(Esc $h[1])"" display=""$(Esc $h[2])""/>") }
        [void]$script:Buf.Append('</hyperlinks>')
    }
    [void]$script:Buf.Append('<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
    [void]$script:Buf.Append('<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
    [void]$script:Buf.Append("<headerFooter><oddFooter>&amp;C&amp;P / &amp;N</oddFooter></headerFooter></worksheet>`n")
    $parts["xl/worksheets/sheet$($script:SheetIdx).xml"] = $script:Buf.ToString()
}
function Get-Cond($B, [string]$key) {
    $v = ''
    foreach ($c in $B.Conds) { if ($c[0] -ceq $key) { $v = $c[1] } }
    $v.Replace("`n", '   ')
}
function Get-LevelStyle([string]$v, [int]$band) {
    if ($reLvErr.IsMatch($v)) { return $S_ERR }
    if ($v -ceq 'WARN') { return $S_WRN }
    if ($v -cmatch '^(DEBUG|TRACE)$') { return $S_DIM + $band }
    $S_CTR + $band
}
function Test-Positive([string]$v) { $reNum.IsMatch($v) -and [double]$v -gt 0 }

# 明細シート
function Write-Detail([int]$bi) {
    $B = $blocks[$bi - 1]; $t = $B.Type
    $defs = $COLS[$t].Split('|'); $nc = $defs.Count
    $HD = @{}; $KD = @{}; $MN = @{}; $MX = @{}; $W = @{}
    for ($c = 1; $c -le $nc; $c++) {
        $y = $defs[$c - 1].Split(':'); $HD[$c] = $y[0]; $KD[$c] = $y[1]
        $len = [math]::Floor((Get-Width $y[0]) * 1.2 + 5.5)        # 見出し (太字) + フィルタボタンが収まる幅
        $MN[$c] = [math]::Max([int]$y[2], $len); $MX[$c] = [math]::Max([int]$y[3], $MN[$c]); $W[$c] = $MN[$c]
    }
    for ($c = 1; $c -le $nc; $c++) {                 # 列幅 (Get-Width と同じ計算を行数分インラインで実施)
        if ($MX[$c] -le $MN[$c]) { continue }
        $wc = $W[$c]; $cap = $MX[$c]
        foreach ($r in $B.Rows) {
            if ($wc -ge $cap) { break }
            $v = [string]$r[$c]
            if (2 * $v.Length + 2 -le $wc) { continue }
            $i = $v.IndexOf("`n"); if ($i -ge 0) { $v = $v.Substring(0, $i) }
            $len = 2 + $(if ($v.Length -gt 200) { 200 } else { $v.Length + $reNarrow.Replace($v, '').Length })
            if ($len -gt $wc) { $wc = $len }
        }
        $W[$c] = [math]::Min($wc, $cap)
    }
    Sheet-Open ($bi + 1) 4 $nc $W
    $last = $CLN[$nc]; $n = $B.Rows.Count
    Row-Open 1 30; Cell-Str 1 "No.$bi  JBoss EAP $t 検索結果" $S_TITLE; for ($c = 2; $c -le $nc; $c++) { Cell-Str $c '' $S_TITLE }; Row-Close
    Row-Open 2 18; Cell-Str 1 ('期間 (JST) : ' + (Get-Cond $B '期間 (JST)') + '      対象 : ' + (Get-Cond $B '対象')) $S_META; Row-Close
    Row-Open 3 18; Cell-Str 1 ('絞り込み : ' + (Get-Cond $B '絞り込み') + "      件数 : $n 件 (ERROR $($B.EC) / WARN $($B.WC))      入力ログ : $($B.Src)") $S_META; Row-Close
    Row-Open 4 24; for ($c = 1; $c -le $nc; $c++) { Cell-Str $c $HD[$c] $S_HEAD }; Row-Close
    # 明細行は件数が多いため、Cell-* 関数と同じ XML をループ内で直接組み立てる (PowerShell の関数呼び出しは低速)
    $KS = @{ dtms = $S_DTMS; dts = $S_DTS; code = $S_CODE; int = $S_INT; dec = $S_DEC; ctr = $S_CTR; txt = $S_TEXT; lvl = $S_CTR }
    $buf = $script:Buf; $d = [datetime]::MinValue; $k = 0
    foreach ($r in $B.Rows) {
        $k++; $band = ($k - 1) % 2; $rn = $k + 4
        [void]$buf.Append('<row r="').Append($rn).Append('" ht="18" customHeight="1"><c r="A').Append($rn).Append('" s="').Append($S_CODE + $band).Append('"><v>').Append($k).Append('</v></c>')
        for ($c = 2; $c -le $nc; $c++) {
            $v = [string]$r[$c]; $kind = $KD[$c]; $st = $KS[$kind] + $band; $num = ''
            if ($kind -eq 'lvl') {                                  # Get-LevelStyle と同じ判定
                $st = if ($reLvErr.IsMatch($v)) { $S_ERR } elseif ($v -ceq 'WARN') { $S_WRN } elseif ($v -ceq 'DEBUG' -or $v -ceq 'TRACE') { $S_DIM + $band } else { $S_CTR + $band }
            } elseif ($kind -eq 'dtms' -or $kind -eq 'dts') {         # ConvertTo-Serial と同じ変換
                $m = $reDate.Match($v)
                if ($m.Success -and [datetime]::TryParseExact($m.Groups[1].Value, 'yyyy-MM-dd HH:mm:ss', $INV, [Globalization.DateTimeStyles]::None, [ref]$d)) {
                    if ($m.Groups[2].Success) { $d = $d.AddMilliseconds([int]$m.Groups[2].Value.Substring(1)) }
                    $num = $d.ToOADate().ToString('0.0000000000', $INV)
                } else { $st = $S_CTR + $band }
            } elseif (($kind -eq 'code' -or $kind -eq 'int' -or $kind -eq 'dec') -and $reNum.IsMatch($v)) {
                $num = $v
            }
            [void]$buf.Append('<c r="').Append($CLN[$c]).Append($rn).Append('" s="').Append($st)
            if ($num) { [void]$buf.Append('"><v>').Append($num).Append('</v></c>') }
            elseif ($v -eq '') { [void]$buf.Append('"/>') }
            else {
                if ($v.Length -gt 32000) { $v = $v.Substring(0, 32000) + "`n...(32,000 文字を超えるため省略)" }
                [void]$buf.Append('" t="inlineStr"><is><t xml:space="preserve">').Append($reCtl.Replace($v, '').Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')).Append('</t></is></c>')
            }
        }
        [void]$buf.Append("</row>`n")
    }
    if ($n -eq 0) { Row-Open 5 18; Cell-Str 1 $(if ($B.Err) { $B.Err } else { '該当するログはありません' }) $S_META; Row-Close }
    Sheet-Close "A4:$last$($n + 4)"
    $B.FilterRange = "'$($B.Sheet)'!`$A`$4:`$$last`$$($n + 4)"
}

# サマリー シート (A:余白 B:見出し C～J:数値 K:長い文字列)
function Add-KV([string]$lab, [string]$val) {
    $n = $val.Split("`n").Count
    Row-Open $script:R_ $(if ($n -gt 1) { 18 * $n } else { 20 }); Cell-Str 2 $lab $S_LABEL; Cell-Merge 3 11 $val $S_VALUE; Row-Close; $script:R_++
}
function Add-Section([string]$txt) { Row-Open $script:R_ 26; Cell-Str 2 $txt $S_SEC; for ($c = 3; $c -le 11; $c++) { Cell-Str $c '' $S_SEC }; Row-Close; $script:R_++ }
function Add-Caption([string]$txt) { $script:R_++; Row-Open $script:R_ 20; Cell-Str 2 $txt $S_CAP; Row-Close; $script:R_++ }
function Add-Header([string]$list) {                 # "見出し:開始列:終了列|..."
    Row-Open $script:R_ 22
    foreach ($x in $list.Split('|')) {
        $a = $x.Split(':'); $c1 = [int]$a[1]; $c2 = [int]$a[2]
        if ($c2 -gt $c1) { Cell-Merge $c1 $c2 $a[0] $S_HEAD } else { Cell-Str $c1 $a[0] $S_HEAD }
    }
    Row-Close; $script:R_++
}
function Write-Summary {
    $W = @{ 1 = 2; 2 = 30; 11 = 60 }; for ($c = 3; $c -le 10; $c++) { $W[$c] = 11 }
    Sheet-Open 1 0 11 $W
    Row-Open 2 34; Cell-Str 2 'JBoss EAP ログ検索レポート' $S_TITLE; for ($c = 3; $c -le 11; $c++) { Cell-Str $c '' $S_TITLE }; Row-Close
    $script:R_ = 4
    Add-KV '作成日時' $NOW
    foreach ($f in $inputs) { Add-KV '入力ログ' $f }
    Add-KV '検索数' "$($blocks.Count) 件"
    $script:R_++; Add-Section '検索一覧'
    Add-Header '明細シート:2:2|HITS:3:3|ERROR:4:4|WARN:5:5|ファイル数:6:6|ログ種別:7:7|期間 (JST):8:11'
    for ($bi = 1; $bi -le $blocks.Count; $bi++) {
        $B = $blocks[$bi - 1]; $band = ($bi - 1) % 2
        Row-Open $script:R_ 20
        Cell-Link 2 $B.Sheet ($S_LINK + $band); Cell-Num 3 $B.Rows.Count ($S_INT + $band)
        Cell-Num 4 $B.EC ($(if ($B.EC) { $S_INTB } else { $S_INT }) + $band); Cell-Num 5 $B.WC ($S_INT + $band)
        Cell-Num 6 $B.FileCount ($S_INT + $band); Cell-Str 7 $B.Type ($S_CTR + $band); Cell-Merge 8 11 (Get-Cond $B '期間 (JST)') ($S_TEXT + $band)
        Row-Close; $script:R_++
    }
    for ($bi = 1; $bi -le $blocks.Count; $bi++) {
        $B = $blocks[$bi - 1]; $gc = $B.Type -eq 'gc.log'; $ac = $B.Type -eq 'access_log'
        $script:R_++; Add-Section "No.$bi  $($B.Type) 検索結果"
        Row-Open $script:R_ 20; Cell-Str 2 '明細シート' $S_LABEL; Cell-Link 3 $B.Sheet $S_LINK; for ($c = 4; $c -le 11; $c++) { Cell-Str $c '' $S_LINK }
        $script:MG.Add("C$($script:R_):K$($script:R_)"); Row-Close; $script:R_++
        Add-KV '入力ログ' $B.Src
        foreach ($c in $B.Conds) { Add-KV $c[0] $c[1] }
        Add-KV '件数' "$($B.Rows.Count) 件  (ERROR $($B.EC) / WARN $($B.WC))"
        if ($B.Err) { Add-KV 'エラー' $B.Err }
        foreach ($x in $B.Notes) { Add-KV '注意' $x }
        if ($B.Groups.Count -or $B.Total) {
            Add-Caption '■ CONTAINER / SERVICE 別'
            Add-Header ('CONTAINER / SERVICE:2:2|HITS:3:3|ERROR:4:4|WARN:5:5' + $(if ($gc) { '|PAUSES:6:6|FULL:7:7|MAX(ms):8:8|TOTAL(ms):9:9' }))
            $k = 0
            foreach ($g in $B.Groups) {
                $band = $k % 2; $k++
                Row-Open $script:R_ 20; Cell-Str 2 $g[1] ($S_TEXT + $band)
                for ($c = 2; $c -lt $g.Count; $c++) {
                    $sty = if ($c -ge 7) { $S_DEC } elseif (($c -eq 3 -or $c -eq 6) -and (Test-Positive $g[$c])) { $S_INTB } else { $S_INT }
                    Cell-Num ($c + 1) $g[$c] ($sty + $band)
                }
                Row-Close; $script:R_++
            }
            if ($B.Total) {
                Row-Open $script:R_ 20; Cell-Str 2 'TOTAL' $S_TTEXT
                for ($c = 2; $c -le $B.Total.Count; $c++) { Cell-Num ($c + 1) $B.Total[$c - 1] $(if ($c -ge 7) { $S_TDEC } else { $S_TINT }) }
                Row-Close; $script:R_++
            }
        }
        if ($ac -and $B.Status.Count) {
            Add-Caption '■ STATUS 別'
            Add-Header 'STATUS:2:2|件数:3:3'
            $k = 0
            foreach ($x in $B.Status) { $band = $k % 2; $k++; Row-Open $script:R_ 20; Cell-Num 2 $x[0] ($S_CODE + $band); Cell-Num 3 $x[1] ($S_INT + $band); Row-Close; $script:R_++ }
        }
        if ($gc -and $B.Top.Count) {
            Add-Caption "■ 停止時間 TOP $($B.Top.Count)"
            Add-Header 'TIMESTAMP (JST):2:2|CTR:3:3|SERVICE:4:4|TASK:5:5|PAUSE(ms):6:6|GC LOG:7:11'
            $k = 0
            foreach ($x in $B.Top) {
                $band = $k % 2; $k++; Row-Open $script:R_ 20
                Cell-Date 2 $x[0] ($S_DTMS + $band) ($S_CTR + $band); Cell-Str 3 $x[1] ($S_CTR + $band); Cell-Str 4 $x[2] ($S_TEXT + $band)
                Cell-Str 5 $x[3] ($S_CTR + $band); Cell-Num 6 $x[4] ($S_DEC + $band); Cell-Merge 7 11 $x[5] ($S_TEXT + $band)
                Row-Close; $script:R_++
            }
        }
        if ($B.Files.Count) {
            Add-Caption "■ ファイル別 ($($B.FileCount) 件)"
            $f = 6 + [int]$gc + [int]($ac -or $gc)   # FILE 列の開始位置
            Add-Header ('CONTAINER / SERVICE / TASK:2:2|HITS:3:3|ERROR:4:4|WARN:5:5' + $(if ($gc) { '|MAX(ms):6:6' }) + $(if ($ac -or $gc) { "|解析不可行:$($f - 1):$($f - 1)" }) + "|FILE:${f}:11")
            $k = 0
            foreach ($x in $B.Files) {
                $band = $k % 2; $k++; Row-Open $script:R_ 20
                $sp = $x[5].Split('/'); $n = $sp.Count
                $who = if ($n -ge 6) { "$($sp[$n - 6]) / $($sp[$n - 4]) / " + $(if ($sp[$n - 2].Length -eq 32) { $sp[$n - 2].Substring(0, 8) } else { $sp[$n - 2] }) } else { '' }
                Cell-Str 2 $who ($S_TEXT + $band)
                Cell-Num 3 $x[0] ($S_INT + $band); Cell-Num 4 $x[1] ($(if (Test-Positive $x[1]) { $S_INTB } else { $S_INT }) + $band); Cell-Num 5 $x[2] ($S_INT + $band)
                if ($gc) { Cell-Num 6 $x[3] ($S_DEC + $band) }
                if ($ac -or $gc) { Cell-Num ($f - 1) $x[4] ($S_INT + $band) }
                Cell-Merge $f 11 $x[5] ($S_TEXT + $band)
                Row-Close; $script:R_++
            }
        }
    }
    Sheet-Close ''
}

$fmt = if ($blocks.Count -gt 99) { '{0:000}_{1}' } else { '{0:00}_{1}' }
for ($bi = 1; $bi -le $blocks.Count; $bi++) { $B = $blocks[$bi - 1]; $B.Sheet = $fmt -f $bi, $B.Type; $B.FilterRange = '' }
for ($bi = 1; $bi -le $blocks.Count; $bi++) { Write-Detail $bi }
Write-Summary

$nb = $blocks.Count
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("$XMLH<workbook xmlns=""$NS"" xmlns:r=""$NSR""><bookViews><workbookView activeTab=""0""/></bookViews><sheets>")
[void]$sb.Append('<sheet name="サマリー" sheetId="1" r:id="rId1"/>')
for ($bi = 1; $bi -le $nb; $bi++) { [void]$sb.Append("<sheet name=""$(Esc $blocks[$bi - 1].Sheet)"" sheetId=""$($bi + 1)"" r:id=""rId$($bi + 1)""/>") }
[void]$sb.Append('</sheets><definedNames>')
for ($bi = 1; $bi -le $nb; $bi++) { [void]$sb.Append("<definedName name=""_xlnm._FilterDatabase"" localSheetId=""$bi"" hidden=""1"">$(Esc $blocks[$bi - 1].FilterRange)</definedName>") }
for ($bi = 1; $bi -le $nb; $bi++) { [void]$sb.Append("<definedName name=""_xlnm.Print_Titles"" localSheetId=""$bi"">$(Esc ("'" + $blocks[$bi - 1].Sheet + "'!`$4:`$4"))</definedName>") }
[void]$sb.Append("</definedNames></workbook>`n")
$parts['xl/workbook.xml'] = $sb.ToString()

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("$XMLH<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">")
for ($i = 1; $i -le $nb + 1; $i++) { [void]$sb.Append("<Relationship Id=""rId$i"" Type=""$NSR/worksheet"" Target=""worksheets/sheet$i.xml""/>") }
[void]$sb.Append("<Relationship Id=""rId$($nb + 2)"" Type=""$NSR/styles"" Target=""styles.xml""/></Relationships>`n")
$parts['xl/_rels/workbook.xml.rels'] = $sb.ToString()

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("$XMLH<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">")
[void]$sb.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>')
[void]$sb.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
[void]$sb.Append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
for ($i = 1; $i -le $nb + 1; $i++) { [void]$sb.Append("<Override PartName=""/xl/worksheets/sheet$i.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>") }
[void]$sb.Append("<Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml""/></Types>`n")
$parts['[Content_Types].xml'] = $sb.ToString()

$parts['_rels/.rels'] = "$XMLH<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">" +
    "<Relationship Id=""rId1"" Type=""$NSR/officeDocument"" Target=""xl/workbook.xml""/>" +
    "<Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/></Relationships>`n"

$parts['docProps/core.xml'] = "$XMLH<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" +
    '<dc:title>JBoss EAP ログ検索レポート</dc:title><dc:creator>jboss_log_report</dc:creator>' +
    "<dcterms:created xsi:type=""dcterms:W3CDTF"">$NOWZ</dcterms:created><dcterms:modified xsi:type=""dcterms:W3CDTF"">$NOWZ</dcterms:modified></cp:coreProperties>`n"

# スタイル : Meiryo UI / モノトーン (見出し=濃灰 / 縞=淡灰 / ERROR=黒地白字 / WARN=灰地太字 / DEBUG,TRACE=灰字)
$FONT = '<name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>'
$BD   = '<left style="thin"><color rgb="{0}"/></left><right style="thin"><color rgb="{0}"/></right><top style="thin"><color rgb="{0}"/></top><bottom style="thin"><color rgb="{0}"/></bottom><diagonal/>'
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("$XMLH<styleSheet xmlns=""$NS"">")
[void]$sb.Append('<numFmts count="3"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss.000"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/><numFmt numFmtId="166" formatCode="#,##0.000"/></numFmts>')
[void]$sb.Append('<fonts count="8">')
[void]$sb.Append('<font><sz val="10"/><color rgb="FF262626"/>' + $FONT)            # 0 標準
[void]$sb.Append('<font><b/><sz val="10"/><color rgb="FF262626"/>' + $FONT)        # 1 太字
[void]$sb.Append('<font><b/><sz val="16"/><color rgb="FF000000"/>' + $FONT)        # 2 タイトル
[void]$sb.Append('<font><b/><sz val="10"/><color rgb="FFFFFFFF"/>' + $FONT)        # 3 白太字
[void]$sb.Append('<font><sz val="9"/><color rgb="FF595959"/>' + $FONT)             # 4 補足
[void]$sb.Append('<font><sz val="10"/><color rgb="FF8C8C8C"/>' + $FONT)            # 5 淡色
[void]$sb.Append('<font><u/><sz val="10"/><color rgb="FF262626"/>' + $FONT)        # 6 リンク
[void]$sb.Append('<font><b/><sz val="12"/><color rgb="FF000000"/>' + $FONT)        # 7 セクション
[void]$sb.Append('</fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>')
foreach ($rgb in 'FF404040', 'FFF2F2F2', 'FF000000', 'FFBFBFBF', 'FFE4E4E4') {     # 2 見出し 3 縞 4 ERROR 5 WARN 6 ラベル
    [void]$sb.Append("<fill><patternFill patternType=""solid""><fgColor rgb=""$rgb""/><bgColor indexed=""64""/></patternFill></fill>")
}
[void]$sb.Append('</fills><borders count="4"><border><left/><right/><top/><bottom/><diagonal/></border>')
[void]$sb.Append('<border>' + ($BD -f 'FFBFBFBF') + '</border>')                    # 1 罫線 (淡)
[void]$sb.Append('<border>' + ($BD -f 'FF808080') + '</border>')                    # 2 罫線 (見出し)
[void]$sb.Append('<border><left/><right/><top/><bottom style="medium"><color rgb="FF404040"/></bottom><diagonal/></border></borders>')
[void]$sb.Append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
# numFmt font fill border 横位置 折返し
$xfs = ('0 0 0 0 - 0|0 2 0 3 - 0|0 4 0 0 - 0|0 7 0 3 - 0|0 1 0 0 - 0|0 3 2 2 center 1|0 1 6 1 left 0|0 0 0 1 left 1|' +
        '0 0 0 1 - 0|0 0 3 1 - 0|0 0 0 1 center 0|0 0 3 1 center 0|164 0 0 1 center 0|164 0 3 1 center 0|165 0 0 1 center 0|165 0 3 1 center 0|' +
        '3 0 0 1 right 0|3 0 3 1 right 0|166 0 0 1 right 0|166 0 3 1 right 0|1 0 0 1 center 0|1 0 3 1 center 0|0 6 0 1 - 0|0 6 3 1 - 0|' +
        '0 5 0 1 center 0|0 5 3 1 center 0|0 3 4 1 center 0|0 1 5 1 center 0|0 1 6 1 - 0|3 1 6 1 right 0|166 1 6 1 right 0|3 1 0 1 right 0|3 1 3 1 right 0').Split('|')
[void]$sb.Append("<cellXfs count=""$($xfs.Count)"">")
foreach ($x in $xfs) {
    $a = $x.Split(' ')
    $h = if ($a[4] -eq '-') { '' } else { " horizontal=""$($a[4])""" }
    $wr = if ($a[5] -eq '1') { ' wrapText="1"' } else { '' }
    [void]$sb.Append("<xf numFmtId=""$($a[0])"" fontId=""$($a[1])"" fillId=""$($a[2])"" borderId=""$($a[3])"" xfId=""0"" applyNumberFormat=""1"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1""><alignment$h vertical=""center""$wr/></xf>")
}
[void]$sb.Append("</cellXfs><cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles></styleSheet>`n")
$parts['xl/styles.xml'] = $sb.ToString()

# --- xlsx (zip) にまとめる ---
Add-Type -AssemblyName System.IO.Compression
$utf8 = New-Object System.Text.UTF8Encoding($false)
$order = @('[Content_Types].xml', '_rels/.rels', 'docProps/core.xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml') +
         @(1..($nb + 1) | ForEach-Object { "xl/worksheets/sheet$_.xml" })
try {
    $fs = [IO.File]::Open($OutFile, [IO.FileMode]::Create)
} catch {
    Fail "出力ファイルを作成できません (Excel で開いていないか確認してください) : $OutFile"
}
try {
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    foreach ($name in $order) {
        $w = New-Object System.IO.StreamWriter($zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal).Open(), $utf8)
        $w.Write($parts[$name]); $w.Dispose()
    }
    $zip.Dispose()
} finally {
    $fs.Dispose()
}

$total = 0; foreach ($B in $blocks) { $total += $B.Rows.Count }
Write-Host ('検索結果 {0} 件 / 明細 {1} 行' -f $nb, $total)
foreach ($B in $blocks) { Write-Host ('  {0,-16} {1,7} 件  (ERROR {2} / WARN {3})' -f $B.Sheet, $B.Rows.Count, $B.EC, $B.WC) }
Write-Host "[OK] Excel レポートを出力しました : $OutFile"
