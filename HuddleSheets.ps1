#requires -Version 5.1
<#
    SST & Shift Huddle - weekly sign-in sheet builder

    Run "Huddle Sheets.bat" (double-click, or drag the email / attachments onto it).
    Type 'help' at the prompt for every command.

    Files next to this script:
      HuddleConfig.xlsx       shops, shifts, employees, settings (edit here or with commands)
      SignSheetTemplate.xlsx  layout of one sign-in sheet (edit freely; I5 must stay merged
                              down over the employee rows)
      lib\PdfSharp.dll        used to combine PDFs (MIT licensed, pdfsharp.net)
#>
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$CmdArgs
)

$ErrorActionPreference = 'Stop'

$Script:Root         = $PSScriptRoot
$Script:ConfigPath   = Join-Path $Script:Root 'HuddleConfig.xlsx'
$Script:TemplatePath = Join-Path $Script:Root 'SignSheetTemplate.xlsx'
$Script:PdfSharpPath = Join-Path (Join-Path $Script:Root 'lib') 'PdfSharp.dll'
$Script:Inv          = [Globalization.CultureInfo]::InvariantCulture
$Script:US           = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$Script:Latin1       = [Text.Encoding]::GetEncoding(28591)
$Script:DayNames     = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')
$Script:DayAbbr      = @('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat')
$Script:WordExt      = @('.docx', '.doc', '.docm', '.dotx', '.rtf', '.odt', '.txt')
$Script:PptExt       = @('.pptx', '.ppt', '.pptm', '.ppsx', '.pps', '.odp')
$Script:XlExt        = @('.xlsx', '.xls', '.xlsm', '.ods')
$Script:ImgExt       = @('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff')
$Script:TopicExt     = @('.pdf') + $Script:WordExt + $Script:PptExt + $Script:XlExt + $Script:ImgExt
$Script:PdfSharpOK   = $false

$Script:DefaultSettings = [ordered]@{
    OutputFolder   = @('',             'Where PDFs are saved. Blank = "PDFs" folder next to the script. Can be a full path.')
    ExtraBlankRows = @('3',            'Blank rows printed under the names on each sheet (fill-ins / new people).')
    MinRows        = @('10',           'Minimum number of name rows on a sheet.')
    NightShift     = @('3',            'Shift that starts the evening before - its dates print 1 day earlier (Mon column = Sunday date). 0 = none.')
    Shift1Name     = @('First Shift',  'Shift label printed on the sheet.')
    Shift2Name     = @('Second Shift', 'Shift label printed on the sheet.')
    Shift3Name     = @('Third Shift',  'Shift label printed on the sheet.')
    TopicCopies    = @('1',            'Default number of safety-topic copies when exporting with topics.')
    OpenWhenDone   = @('Yes',          'Open the PDF when it is done (Yes/No).')
    OutlookSubject = @('Safety Start Up Talks', 'Text in the subject of the weekly topics email (for the [O] Outlook option).')
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

#region ---------- console helpers ----------

function Say([string]$Text = '', [string]$Color = 'Gray') { Write-Host $Text -ForegroundColor $Color }
function Head([string]$Text) { Write-Host ''; Write-Host $Text -ForegroundColor Cyan }

# Prompt; 'cancel' at any prompt backs out of the current command.
function Ask([string]$Prompt, [string]$Default = '') {
    $p = $Prompt
    if ($Default -ne '') { $p += " [$Default]" }
    $a = Read-Host $p
    if ($null -eq $a) { $a = '' }
    $a = $a.Trim()
    if ($a -ieq 'cancel') { throw (New-Object OperationCanceledException) }
    if ($a -eq '') { return $Default }
    return $a
}

function AskYesNo([string]$Prompt, [bool]$Default = $false) {
    $d = 'y/N'
    if ($Default) { $d = 'Y/n' }
    while ($true) {
        $a = (Ask "$Prompt ($d)").ToLower()
        if ($a -eq '') { return $Default }
        if ($a -eq 'y' -or $a -eq 'yes') { return $true }
        if ($a -eq 'n' -or $a -eq 'no') { return $false }
    }
}

function AskInt([string]$Prompt, [int]$Min, [int]$Max, [int]$Default) {
    while ($true) {
        $a = Ask $Prompt "$Default"
        $n = 0
        if ([int]::TryParse($a, [ref]$n) -and $n -ge $Min -and $n -le $Max) { return $n }
        Say "  Enter a number from $Min to $Max." Yellow
    }
}

function Get-Ordinal([int]$n) {
    switch ($n) { 1 { return '1st' } 2 { return '2nd' } 3 { return '3rd' } default { return "${n}th" } }
}

function Clean-PathText([string]$s) {
    $s = "$s".Trim()
    if ($s.StartsWith('& ')) { $s = $s.Substring(2).Trim() }
    return $s.Trim('"').Trim("'").Trim()
}

# Splits a pasted / dragged line into paths ("a b.docx" "c.pdf" or single unquoted path)
function Split-PathList([string]$s) {
    $s = "$s".Trim()
    if ($s.StartsWith('& ')) { $s = $s.Substring(2).Trim() }
    $whole = Clean-PathText $s
    if ($whole -ne '' -and (Test-Path -LiteralPath $whole)) { return @($whole) }
    $out = @()
    foreach ($m in [regex]::Matches($s, '"([^"]+)"|''([^'']+)''|(\S+)')) {
        $v = $m.Groups[1].Value + $m.Groups[2].Value + $m.Groups[3].Value
        if ($v -ne '') { $out += $v }
    }
    return $out
}

function Get-SafeFileName([string]$name) {
    $bad = [IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object Text.StringBuilder
    foreach ($ch in $name.ToCharArray()) { if ($bad -contains $ch) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) } }
    $r = $sb.ToString().Trim().TrimEnd('.')
    if ($r -eq '') { $r = 'file' }
    return $r
}

function Get-UniquePath([string]$Dir, [string]$Name) {
    $p = Join-Path $Dir $Name
    $base = [IO.Path]::GetFileNameWithoutExtension($Name)
    $ext = [IO.Path]::GetExtension($Name)
    $i = 2
    while (Test-Path -LiteralPath $p) { $p = Join-Path $Dir ("{0} ({1}){2}" -f $base, $i, $ext); $i++ }
    return $p
}

# Prints the error with the script line it came from, and appends details to HuddleSheets-error.log
function Show-Error($Err) {
    Say ('  Error: ' + $Err.Exception.Message) Red
    $ln = $Err.InvocationInfo.ScriptLineNumber
    if ($ln) { Say ("  (line {0}: {1})" -f $ln, "$($Err.InvocationInfo.Line)".Trim()) DarkGray }
    try {
        $log = Join-Path $Script:Root 'HuddleSheets-error.log'
        $txt = "[{0}] {1}`r`n{2}`r`n{3}`r`n{4}`r`n" -f (Get-Date), $Err.Exception.ToString(), $Err.InvocationInfo.PositionMessage, $Err.ScriptStackTrace, ('-' * 60)
        Add-Content -LiteralPath $log -Value $txt
        Say "  Details saved to HuddleSheets-error.log" DarkGray
    }
    catch { }
}

function Remove-Com($o) {
    if ($null -ne $o) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } catch { } }
}

#endregion

#region ---------- minimal .xlsx reader / writer (no Excel needed) ----------

function Get-ZipText($Zip, [string]$Name) {
    $e = $Zip.GetEntry($Name)
    if ($null -eq $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open())
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

function Get-ColIndex([string]$Ref) {
    $n = 0
    foreach ($ch in (($Ref -replace '\d', '').ToUpper()).ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n - 1
}

function Get-ColLetter([int]$Index) {
    $s = ''; $i = $Index + 1
    while ($i -gt 0) { $m = ($i - 1) % 26; $s = [string][char](65 + $m) + $s; $i = [math]::Floor(($i - 1) / 26) }
    return $s
}

# Returns ordered dictionary: sheet name -> ArrayList of string[] (one per row, header first)
function Read-Xlsx([string]$Path) {
    $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $zip = $null
    try {
        $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Read)
        $shared = New-Object 'System.Collections.Generic.List[string]'
        $sst = Get-ZipText $zip 'xl/sharedStrings.xml'
        if ($sst) {
            $x = [xml]$sst
            foreach ($si in $x.DocumentElement.SelectNodes("*[local-name()='si']")) {
                $sb = New-Object Text.StringBuilder
                foreach ($t in $si.SelectNodes(".//*[local-name()='t' and not(ancestor::*[local-name()='rPh'])]")) { [void]$sb.Append($t.InnerText) }
                $shared.Add($sb.ToString())
            }
        }
        $wbx = [xml](Get-ZipText $zip 'xl/workbook.xml')
        $relx = [xml](Get-ZipText $zip 'xl/_rels/workbook.xml.rels')
        $rel = @{}
        foreach ($r in $relx.DocumentElement.SelectNodes("*[local-name()='Relationship']")) { $rel[$r.GetAttribute('Id')] = $r.GetAttribute('Target') }
        $result = [ordered]@{}
        foreach ($s in $wbx.SelectNodes("//*[local-name()='sheet']")) {
            $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $target = $rel[$rid]
            if ($target.StartsWith('/')) { $target = $target.TrimStart('/') } else { $target = 'xl/' + $target }
            $sx = [xml](Get-ZipText $zip $target)
            $rows = New-Object System.Collections.ArrayList
            foreach ($row in $sx.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
                $cells = @{}; $max = -1; $auto = 0
                foreach ($c in $row.SelectNodes("*[local-name()='c']")) {
                    $ref = $c.GetAttribute('r')
                    if ($ref) { $ci = Get-ColIndex $ref } else { $ci = $auto }
                    $auto = $ci + 1
                    $t = $c.GetAttribute('t')
                    $v = $c.SelectSingleNode("*[local-name()='v']")
                    $val = ''
                    if ($t -eq 's') { if ($null -ne $v) { $val = $shared[[int]$v.InnerText] } }
                    elseif ($t -eq 'inlineStr') { $is = $c.SelectSingleNode("*[local-name()='is']"); if ($null -ne $is) { $val = $is.InnerText } }
                    elseif ($t -eq 'b') { if ($null -ne $v) { if ($v.InnerText -eq '1') { $val = 'TRUE' } else { $val = 'FALSE' } } }
                    elseif ($null -ne $v) { $val = $v.InnerText }
                    $cells[$ci] = $val
                    if ($ci -gt $max) { $max = $ci }
                }
                $arr = New-Object string[] ($max + 1)
                foreach ($k in $cells.Keys) { $arr[$k] = $cells[$k] }
                [void]$rows.Add($arr)
            }
            $result[$s.GetAttribute('name')] = $rows
        }
        return $result
    }
    finally {
        if ($null -ne $zip) { $zip.Dispose() }
        $fs.Dispose()
    }
}

# First row = headers. Returns ArrayList of hashtables keyed by header (case-insensitive).
function Get-SheetTable($Book, [string]$Sheet) {
    $out = New-Object System.Collections.ArrayList
    $key = $null
    foreach ($k in $Book.Keys) { if ($k -ieq $Sheet) { $key = $k } }
    if ($null -eq $key) { return , $out }
    $rows = $Book[$key]
    if ($rows.Count -lt 1) { return , $out }
    $hdr = $rows[0]
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $o = @{}; $any = $false
        for ($c = 0; $c -lt $hdr.Count; $c++) {
            $h = "$($hdr[$c])".Trim()
            if ($h -eq '') { continue }
            $v = ''
            if ($c -lt $rows[$i].Count) { $v = "$($rows[$i][$c])".Trim() }
            if ($v -ne '') { $any = $true }
            $o[$h] = $v
        }
        if ($any) { [void]$out.Add($o) }
    }
    return , $out
}

function Esc([string]$s) {
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return [Security.SecurityElement]::Escape($s)
}

# $Sheets: array of @{ Name; Columns = @(@{Header;Width;Numeric;Wrap}); Rows = list of object[]; Validations = @(@{Col;Formula}) }
function Write-Xlsx([string]$Path, $Sheets) {
    $nsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $nsRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $parts = [ordered]@{}
    $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    $wbSheets = ''; $wbRels = ''; $defNames = ''
    for ($si = 0; $si -lt $Sheets.Count; $si++) {
        $sh = $Sheets[$si]
        $n = $si + 1
        $ct += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        $wbSheets += "<sheet name=`"$(Esc $sh.Name)`" sheetId=`"$n`" r:id=`"rId$n`"/>"
        $wbRels += "<Relationship Id=`"rId$n`" Type=`"$nsRel/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
        $cols = $sh.Columns
        $lastCol = Get-ColLetter ($cols.Count - 1)
        $rowCount = $sh.Rows.Count + 1
        $sb = New-Object Text.StringBuilder
        [void]$sb.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"$nsMain`" xmlns:r=`"$nsRel`">")
        $sel = ''
        if ($si -eq 0) { $sel = ' tabSelected="1"' }
        [void]$sb.Append("<sheetViews><sheetView workbookViewId=`"0`"$sel><pane ySplit=`"1`" topLeftCell=`"A2`" activePane=`"bottomLeft`" state=`"frozen`"/><selection pane=`"bottomLeft`" activeCell=`"A2`" sqref=`"A2`"/></sheetView></sheetViews>")
        [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/><cols>')
        for ($c = 0; $c -lt $cols.Count; $c++) { [void]$sb.Append("<col min=`"$($c+1)`" max=`"$($c+1)`" width=`"$($cols[$c].Width)`" customWidth=`"1`"/>") }
        [void]$sb.Append('</cols><sheetData><row r="1">')
        for ($c = 0; $c -lt $cols.Count; $c++) { [void]$sb.Append("<c r=`"$(Get-ColLetter $c)1`" t=`"inlineStr`" s=`"1`"><is><t>$(Esc $cols[$c].Header)</t></is></c>") }
        [void]$sb.Append('</row>')
        $r = 2
        foreach ($row in $sh.Rows) {
            [void]$sb.Append("<row r=`"$r`">")
            for ($c = 0; $c -lt $cols.Count; $c++) {
                $v = $null
                if ($c -lt $row.Count) { $v = $row[$c] }
                if ($null -eq $v -or "$v" -eq '') { continue }
                $ref = "$(Get-ColLetter $c)$r"
                $style = ''
                if ($cols[$c].Wrap) { $style = ' s="2"' }
                $num = 0.0
                if ($cols[$c].Numeric -and [double]::TryParse("$v", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$num)) {
                    [void]$sb.Append("<c r=`"$ref`"$style><v>$($num.ToString($Script:Inv))</v></c>")
                }
                else {
                    [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"$style><is><t xml:space=`"preserve`">$(Esc "$v")</t></is></c>")
                }
            }
            [void]$sb.Append('</row>')
            $r++
        }
        [void]$sb.Append('</sheetData>')
        [void]$sb.Append("<autoFilter ref=`"A1:$lastCol$rowCount`"/>")
        $defNames += "<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$si`" hidden=`"1`">'$(Esc $sh.Name)'!`$A`$1:`$$lastCol`$$rowCount</definedName>"
        if ($sh.Validations -and $sh.Validations.Count -gt 0) {
            [void]$sb.Append("<dataValidations count=`"$($sh.Validations.Count)`">")
            foreach ($dv in $sh.Validations) {
                $L = Get-ColLetter $dv.Col
                [void]$sb.Append("<dataValidation type=`"list`" allowBlank=`"1`" showErrorMessage=`"1`" sqref=`"${L}2:${L}2000`"><formula1>$(Esc $dv.Formula)</formula1></dataValidation>")
            }
            [void]$sb.Append('</dataValidations>')
        }
        [void]$sb.Append('<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
        $parts["xl/worksheets/sheet$n.xml"] = $sb.ToString()
    }
    $ct += '</Types>'
    $parts['[Content_Types].xml'] = $ct
    $parts['_rels/.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $stylesRid = $Sheets.Count + 1
    $parts['xl/_rels/workbook.xml.rels'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$wbRels<Relationship Id=`"rId$stylesRid`" Type=`"$nsRel/styles`" Target=`"styles.xml`"/></Relationships>"
    $parts['xl/workbook.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook xmlns=`"$nsMain`" xmlns:r=`"$nsRel`"><bookViews><workbookView/></bookViews><sheets>$wbSheets</sheets><definedNames>$defNames</definedNames></workbook>"
    $parts['xl/styles.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><styleSheet xmlns=`"$nsMain`"><fonts count=`"2`"><font><sz val=`"11`"/><name val=`"Calibri`"/><family val=`"2`"/></font><font><b/><sz val=`"11`"/><color rgb=`"FFFFFFFF`"/><name val=`"Calibri`"/><family val=`"2`"/></font></fonts><fills count=`"3`"><fill><patternFill patternType=`"none`"/></fill><fill><patternFill patternType=`"gray125`"/></fill><fill><patternFill patternType=`"solid`"><fgColor rgb=`"FF1F4E78`"/><bgColor indexed=`"64`"/></patternFill></fill></fills><borders count=`"1`"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=`"1`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`"/></cellStyleXfs><cellXfs count=`"3`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`" xfId=`"0`"/><xf numFmtId=`"0`" fontId=`"1`" fillId=`"2`" borderId=`"0`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`"/><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`" xfId=`"0`" applyAlignment=`"1`"><alignment vertical=`"top`" wrapText=`"1`"/></xf></cellXfs><cellStyles count=`"1`"><cellStyle name=`"Normal`" xfId=`"0`" builtinId=`"0`"/></cellStyles></styleSheet>"

    $dir = Split-Path -Parent $Path
    $tmp = Join-Path $dir ('~' + [IO.Path]::GetFileNameWithoutExtension($Path) + '.tmp')
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
    $utf8 = New-Object Text.UTF8Encoding($false)
    $zip = [IO.Compression.ZipFile]::Open($tmp, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($k in $parts.Keys) {
            $e = $zip.CreateEntry($k)
            $w = New-Object IO.StreamWriter($e.Open(), $utf8)
            try { $w.Write($parts[$k]) } finally { $w.Dispose() }
        }
    }
    finally { $zip.Dispose() }
    try { [IO.File]::Copy($tmp, $Path, $true) }
    catch { throw "Can't save $([IO.Path]::GetFileName($Path)) - it's probably open in Excel. Close it and try again." }
    finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
}

#endregion

#region ---------- config ----------

function New-Employee([string]$Shop, [int]$Shift, [string]$Name, [string]$Assignment) {
    return [pscustomobject]@{ Shop = $Shop.ToUpper(); Shift = $Shift; Name = $Name; Assignment = $Assignment }
}

function Get-Config {
    if (-not (Test-Path -LiteralPath $Script:ConfigPath)) { throw "Config not found: $Script:ConfigPath" }
    $book = Read-Xlsx $Script:ConfigPath
    $cfg = [pscustomobject]@{
        Shops     = New-Object System.Collections.ArrayList
        Employees = New-Object System.Collections.ArrayList
        Settings  = [ordered]@{}
    }
    foreach ($r in (Get-SheetTable $book 'Shops')) {
        $code = "$($r['Code'])".Trim().ToUpper()
        if ($code -eq '') { continue }
        $shifts = 1
        $n = 0.0
        if ([double]::TryParse("$($r['Shifts'])", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$n)) { $shifts = [int]$n }
        if ($shifts -lt 1) { $shifts = 1 }
        $title = "$($r['Title'])".Trim()
        if ($title -eq '') { $title = "$code SST & Shift Huddle" }
        [void]$cfg.Shops.Add([pscustomobject]@{ Code = $code; Title = $title; Shifts = $shifts })
    }
    foreach ($r in (Get-SheetTable $book 'Employees')) {
        $name = "$($r['Name'])".Trim()
        if ($name -eq '') { continue }
        $shift = 1
        $n = 0.0
        if ([double]::TryParse("$($r['Shift'])", [Globalization.NumberStyles]::Float, $Script:Inv, [ref]$n)) { $shift = [int]$n }
        [void]$cfg.Employees.Add((New-Employee "$($r['Shop'])".Trim() $shift $name "$($r['Assignment'])".Trim()))
    }
    foreach ($k in $Script:DefaultSettings.Keys) { $cfg.Settings[$k] = $Script:DefaultSettings[$k][0] }
    foreach ($r in (Get-SheetTable $book 'Settings')) {
        $k = "$($r['Setting'])".Trim()
        if ($k -eq '') { continue }
        $match = $null
        foreach ($dk in @($cfg.Settings.Keys)) { if ($dk -ieq $k) { $match = $dk } }
        if ($null -eq $match) { $match = $k }
        $cfg.Settings[$match] = "$($r['Value'])".Trim()
    }
    return $cfg
}

function Save-Config($Cfg) {
    $order = @{}
    for ($i = 0; $i -lt $Cfg.Shops.Count; $i++) { $order[$Cfg.Shops[$i].Code] = $i }
    $idx = 0
    $tagged = foreach ($e in $Cfg.Employees) {
        $so = 999
        if ($order.ContainsKey($e.Shop)) { $so = $order[$e.Shop] }
        [pscustomobject]@{ E = $e; S = $so; I = $idx }
        $idx++
    }
    $sorted = @($tagged | Sort-Object S, @{ Expression = { $_.E.Shift } }, I | ForEach-Object { $_.E })
    $Cfg.Employees.Clear()
    foreach ($e in $sorted) { [void]$Cfg.Employees.Add($e) }

    $empRows = New-Object System.Collections.ArrayList
    foreach ($e in $Cfg.Employees) { [void]$empRows.Add(@($e.Shop, $e.Shift, $e.Name, $e.Assignment)) }
    $shopRows = New-Object System.Collections.ArrayList
    foreach ($s in $Cfg.Shops) { [void]$shopRows.Add(@($s.Code, $s.Title, $s.Shifts)) }
    $setRows = New-Object System.Collections.ArrayList
    foreach ($k in $Cfg.Settings.Keys) {
        $note = ''
        if ($Script:DefaultSettings.Contains($k)) { $note = $Script:DefaultSettings[$k][1] }
        [void]$setRows.Add(@($k, $Cfg.Settings[$k], $note))
    }
    $sheets = @(
        @{ Name = 'Employees'; Rows = $empRows
           Columns = @(@{ Header = 'Shop'; Width = 9 }, @{ Header = 'Shift'; Width = 7; Numeric = $true }, @{ Header = 'Name'; Width = 32 }, @{ Header = 'Assignment'; Width = 24 })
           Validations = @(@{ Col = 0; Formula = 'Shops!$A$2:$A$60' }, @{ Col = 1; Formula = '"1,2,3"' }) },
        @{ Name = 'Shops'; Rows = $shopRows
           Columns = @(@{ Header = 'Code'; Width = 9 }, @{ Header = 'Title'; Width = 55 }, @{ Header = 'Shifts'; Width = 8; Numeric = $true })
           Validations = @(@{ Col = 2; Formula = '"1,2,3"' }) },
        @{ Name = 'Settings'; Rows = $setRows
           Columns = @(@{ Header = 'Setting'; Width = 18 }, @{ Header = 'Value'; Width = 40 }, @{ Header = 'Notes'; Width = 90; Wrap = $true }) }
    )
    Write-Xlsx $Script:ConfigPath $sheets
}

function Get-Setting($Cfg, [string]$Name) {
    if ($Cfg.Settings.Contains($Name)) { return "$($Cfg.Settings[$Name])" }
    if ($Script:DefaultSettings.Contains($Name)) { return $Script:DefaultSettings[$Name][0] }
    return ''
}

function Get-SettingInt($Cfg, [string]$Name, [int]$Fallback) {
    $n = 0
    if ([int]::TryParse((Get-Setting $Cfg $Name), [ref]$n)) { return $n }
    return $Fallback
}

function Get-ShiftName($Cfg, [int]$Shift) {
    $n = Get-Setting $Cfg "Shift${Shift}Name"
    if ($n -eq '') { $n = (Get-Ordinal $Shift) + ' Shift' }
    return $n
}

function Find-Shop($Cfg, [string]$Text) {
    $t = "$Text".Trim()
    if ($t -eq '') { return $null }
    $n = 0
    if ([int]::TryParse($t, [ref]$n) -and $n -ge 1 -and $n -le $Cfg.Shops.Count) { return $Cfg.Shops[$n - 1] }
    foreach ($s in $Cfg.Shops) { if ($s.Code -ieq $t) { return $s } }
    return $null
}

function Get-ShopEmployees($Cfg, [string]$Code, [int]$Shift) {
    return @($Cfg.Employees | Where-Object { $_.Shop -eq $Code -and $_.Shift -eq $Shift })
}

function Get-OutputRoot($Cfg) {
    $o = Get-Setting $Cfg 'OutputFolder'
    if ($o -eq '') { return (Join-Path $Script:Root 'PDFs') }
    if (-not [IO.Path]::IsPathRooted($o)) { return (Join-Path $Script:Root $o) }
    return $o
}

#endregion

#region ---------- weeks ----------

# Monday of the work week. Sunday belongs to the NEXT week (3rd shift starts Sunday night).
function Get-WeekMonday([datetime]$Date) {
    $d = $Date.Date
    if ($d.DayOfWeek -eq [DayOfWeek]::Sunday) { return $d.AddDays(1) }
    return $d.AddDays( - (([int]$d.DayOfWeek + 6) % 7))
}

function Get-WeekNumber([datetime]$Monday) {
    $thu = $Monday.AddDays(3)                      # ISO-8601 week number
    return [int][math]::Floor(($thu.DayOfYear - 1) / 7) + 1
}

function Get-MondayOfWeek([int]$Week, [int]$Year) {
    $jan4 = New-Object DateTime($Year, 1, 4)
    $wk1 = $jan4.AddDays( - (([int]$jan4.DayOfWeek + 6) % 7))
    return $wk1.AddDays(7 * ($Week - 1))
}

function Get-NearestWeekMonday([int]$Week) {
    $today = (Get-Date).Date
    $best = $null
    foreach ($y in @(($today.Year - 1), $today.Year, ($today.Year + 1))) {
        $m = Get-MondayOfWeek $Week $y
        if ((Get-WeekNumber $m) -ne $Week) { continue }
        if ($null -eq $best -or [math]::Abs(($m - $today).TotalDays) -lt [math]::Abs(($best - $today).TotalDays)) { $best = $m }
    }
    return $best
}

function ConvertTo-DateNear([string]$Text) {
    $t = $Text.Trim() -replace '\.', '/'
    $m = [regex]::Match($t, '^(\d{1,2})[/-](\d{1,2})$')
    $today = (Get-Date).Date
    if ($m.Success) {
        $mo = [int]$m.Groups[1].Value; $da = [int]$m.Groups[2].Value
        $best = $null
        foreach ($y in @(($today.Year - 1), $today.Year, ($today.Year + 1))) {
            try { $d = New-Object DateTime($y, $mo, $da) } catch { continue }
            if ($null -eq $best -or [math]::Abs(($d - $today).TotalDays) -lt [math]::Abs(($best - $today).TotalDays)) { $best = $d }
        }
        return $best
    }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($t, $Script:US, [Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt.Date }
    return $null
}

function Format-Week([datetime]$Monday) {
    return ('Week #{0,-3} Mon {1} - Sat {2}' -f (Get-WeekNumber $Monday), $Monday.ToString('MM/dd/yyyy', $Script:Inv), $Monday.AddDays(5).ToString('MM/dd', $Script:Inv))
}

function Get-WeekFromSubject([string]$Subject) {
    if (-not $Subject) { return $null }
    $m = [regex]::Match($Subject, '(?i)(?:w\s*/\s*o|w\.o\.|week\s+of)\s*:?\s*(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{2,4}))?')
    if (-not $m.Success) { return $null }
    $txt = $m.Groups[1].Value + '/' + $m.Groups[2].Value
    if ($m.Groups[3].Success) { $txt += '/' + $m.Groups[3].Value }
    $d = ConvertTo-DateNear $txt
    if ($null -eq $d) { return $null }
    return (Get-WeekMonday $d)
}

function Select-Week($Suggested) {
    $this = Get-WeekMonday (Get-Date)
    $opts = [ordered]@{ P = $this.AddDays(-7); T = $this; N = $this.AddDays(7) }
    $labels = @{ P = 'Previous week'; T = 'This week'; N = 'Next week' }
    $default = 'T'
    if ($null -ne $Suggested) {
        $hit = $null
        foreach ($k in $opts.Keys) { if ($opts[$k] -eq $Suggested) { $hit = $k } }
        if ($hit) { $default = $hit } else { $opts['E'] = $Suggested; $labels['E'] = 'Email week'; $default = 'E' }
    }
    foreach ($k in $opts.Keys) {
        $mark = ''
        if ($null -ne $Suggested -and $opts[$k] -eq $Suggested) { $mark = '  <- from email' }
        Say ('  [{0}] {1,-14} {2}{3}' -f $k, $labels[$k], (Format-Week $opts[$k]), $mark)
    }
    while ($true) {
        $a = Ask 'Pick P/T/N, type a week # (e.g. 43) or a date (e.g. 10/19)' $default
        $monday = $null
        if ($opts.Contains($a.ToUpper())) { $monday = $opts[$a.ToUpper()] }
        else {
            $wm = [regex]::Match($a, '^(?i)(?:w(?:eek)?\s*)?#?\s*(\d{1,2})$')
            if ($wm.Success) {
                $w = [int]$wm.Groups[1].Value
                if ($w -ge 1 -and $w -le 53) { $monday = Get-NearestWeekMonday $w }
            }
            else {
                $d = ConvertTo-DateNear $a
                if ($null -ne $d) { $monday = Get-WeekMonday $d }
            }
        }
        if ($null -eq $monday) { Say "  Didn't understand '$a'." Yellow; continue }
        Say ("  -> {0}" -f (Format-Week $monday)) Green
        return $monday
    }
}

#endregion

#region ---------- email / attachment intake ----------

function Get-EncodingSafe([string]$Name) {
    $n = "$Name".Trim().Trim('"')
    if ($n -eq '') { return [Text.Encoding]::UTF8 }
    try { return [Text.Encoding]::GetEncoding($n) } catch { return [Text.Encoding]::UTF8 }
}

function ConvertFrom-QPBytes([string]$s) {
    $s = [regex]::Replace($s, '=\r?\n', '')
    $ms = New-Object IO.MemoryStream
    $i = 0
    while ($i -lt $s.Length) {
        $ch = $s[$i]
        if ($ch -eq [char]'=' -and ($i + 2) -lt $s.Length -and $s.Substring($i + 1, 2) -match '^[0-9A-Fa-f]{2}$') {
            $ms.WriteByte([Convert]::ToByte($s.Substring($i + 1, 2), 16)); $i += 3
        }
        else { $ms.WriteByte([byte]([int]$ch -band 0xFF)); $i++ }
    }
    return , $ms.ToArray()
}

function ConvertFrom-PercentBytes([string]$s) {
    $ms = New-Object IO.MemoryStream
    $i = 0
    while ($i -lt $s.Length) {
        if ($s[$i] -eq [char]'%' -and ($i + 2) -lt $s.Length -and $s.Substring($i + 1, 2) -match '^[0-9A-Fa-f]{2}$') {
            $ms.WriteByte([Convert]::ToByte($s.Substring($i + 1, 2), 16)); $i += 3
        }
        else { $ms.WriteByte([byte]([int]$s[$i] -band 0xFF)); $i++ }
    }
    return , $ms.ToArray()
}

function ConvertFrom-EncodedWords([string]$s) {
    if (-not $s) { return $s }
    $s = [regex]::Replace($s, '(\?=)\s+(=\?)', '$1$2')
    $eval = {
        param($m)
        try {
            $enc = Get-EncodingSafe $m.Groups[1].Value
            if ($m.Groups[2].Value -ieq 'B') { $bytes = [Convert]::FromBase64String($m.Groups[3].Value) }
            else { $bytes = ConvertFrom-QPBytes ($m.Groups[3].Value -replace '_', ' ') }
            return $enc.GetString($bytes)
        }
        catch { return $m.Value }
    }
    return [regex]::Replace($s, '=\?([^?]+)\?([BbQq])\?([^?]*)\?=', [Text.RegularExpressions.MatchEvaluator]$eval)
}

# Header parameter incl. RFC 2231 (name*=utf-8''..., name*0*=...) and RFC 2047 words
function Get-HeaderParam([string]$Value, [string]$Name) {
    if (-not $Value) { return $null }
    $n = [regex]::Escape($Name)
    $segs = [regex]::Matches($Value, '(?i)(?:^|;)\s*' + $n + '\*(\d+)(\*?)\s*=\s*(?:"([^"]*)"|([^;]*))')
    if ($segs.Count -gt 0) {
        $list = @($segs | Sort-Object { [int]$_.Groups[1].Value })
        $charset = 'utf-8'
        $ms = New-Object IO.MemoryStream
        foreach ($s in $list) {
            $v = $s.Groups[3].Value + $s.Groups[4].Value.Trim()
            if ($s.Groups[2].Value -eq '*') {
                if ($s.Groups[1].Value -eq '0') {
                    $q = $v.Split([char]39)
                    if ($q.Count -ge 3) { $charset = $q[0]; $v = ($q[2..($q.Count - 1)] -join "'") }
                }
                $b = ConvertFrom-PercentBytes $v
            }
            else { $b = $Script:Latin1.GetBytes($v) }
            $ms.Write($b, 0, $b.Length)
        }
        return (Get-EncodingSafe $charset).GetString($ms.ToArray())
    }
    $m = [regex]::Match($Value, '(?i)(?:^|;)\s*' + $n + '\*\s*=\s*([^;]*)')
    if ($m.Success) {
        $v = $m.Groups[1].Value.Trim().Trim('"')
        $q = $v.Split([char]39)
        if ($q.Count -ge 3) { return (Get-EncodingSafe $q[0]).GetString((ConvertFrom-PercentBytes ($q[2..($q.Count - 1)] -join "'"))) }
        return $v
    }
    $m = [regex]::Match($Value, '(?i)(?:^|;)\s*' + $n + '\s*=\s*(?:"((?:[^"\\]|\\.)*)"|([^;]*))')
    if ($m.Success) {
        $v = ($m.Groups[1].Value + $m.Groups[2].Value.Trim()) -replace '\\(.)', '$1'
        return (ConvertFrom-EncodedWords $v)
    }
    return $null
}

function Split-MimeEntity([string]$Text) {
    $hdrText = ''; $body = ''
    if ($Text -match '^\r?\n') { $body = $Text -replace '^\r?\n', '' }
    else {
        $m = [regex]::Match($Text, '\r?\n\r?\n')
        if ($m.Success) { $hdrText = $Text.Substring(0, $m.Index); $body = $Text.Substring($m.Index + $m.Length) }
        else { $hdrText = $Text }
    }
    $hdrText = [regex]::Replace($hdrText, '\r?\n[ \t]+', ' ')
    $headers = @{}
    foreach ($line in ($hdrText -split '\r?\n')) {
        $i = $line.IndexOf(':')
        if ($i -gt 0) {
            $k = $line.Substring(0, $i).Trim().ToLower()
            if (-not $headers.ContainsKey($k)) { $headers[$k] = $line.Substring($i + 1).Trim() }
        }
    }
    return [pscustomobject]@{ Headers = $headers; Body = $body }
}

function Get-MimeParts([string]$Text, $Acc, [int]$Depth = 0) {
    $ent = Split-MimeEntity $Text
    $ct = $ent.Headers['content-type']
    if (-not $ct) { $ct = 'text/plain' }
    $type = ($ct -split ';')[0].Trim().ToLower()
    if ($type.StartsWith('multipart/')) {
        $b = Get-HeaderParam $ct 'boundary'
        if (-not $b) { return }
        $chunks = [regex]::Split("`r`n" + $ent.Body, '\r?\n--' + [regex]::Escape($b))
        for ($i = 1; $i -lt $chunks.Count; $i++) {
            $ch = $chunks[$i]
            if ($ch.StartsWith('--')) { break }
            $nl = $ch.IndexOf("`n")
            if ($nl -lt 0) { continue }
            Get-MimeParts $ch.Substring($nl + 1) $Acc ($Depth + 1)
        }
        return
    }
    $cte = "$($ent.Headers['content-transfer-encoding'])".Trim().ToLower()
    if ($cte -eq 'base64') {
        $clean = ($ent.Body -replace '[^A-Za-z0-9+/=]', '').TrimEnd('=')
        $clean += ('=' * ((4 - $clean.Length % 4) % 4))
        try { $bytes = [Convert]::FromBase64String($clean) } catch { $bytes = $Script:Latin1.GetBytes($ent.Body) }
    }
    elseif ($cte -eq 'quoted-printable') { $bytes = ConvertFrom-QPBytes $ent.Body }
    else { $bytes = $Script:Latin1.GetBytes($ent.Body) }

    if ($type -eq 'message/rfc822' -and $Depth -lt 6) { Get-MimeParts ($Script:Latin1.GetString($bytes)) $Acc ($Depth + 1); return }

    $disp = "$($ent.Headers['content-disposition'])"
    $fname = Get-HeaderParam $disp 'filename'
    if (-not $fname) { $fname = Get-HeaderParam $ct 'name' }
    $inline = ($ent.Headers.ContainsKey('content-id') -and $disp -notmatch '^\s*attachment')
    [void]$Acc.Add([pscustomobject]@{ Type = $type; FileName = $fname; Bytes = $bytes; Charset = (Get-HeaderParam $ct 'charset'); Inline = $inline })
}

function ConvertFrom-Html([string]$Html) {
    $t = [regex]::Replace($Html, '(?is)<(style|script|head)[^>]*>.*?</\1>', ' ')
    $t = [regex]::Replace($t, '(?i)<br\s*/?>|</p>|</div>|</tr>|</li>|</h\d>', "`n")
    $t = [regex]::Replace($t, '(?i)</td>', ' ')
    $t = [regex]::Replace($t, '<[^>]+>', '')
    return [Net.WebUtility]::HtmlDecode($t)
}

function Read-EmlFile([string]$Path, [string]$Dest) {
    $raw = $Script:Latin1.GetString([IO.File]::ReadAllBytes($Path))
    $top = Split-MimeEntity $raw
    $parts = New-Object System.Collections.ArrayList
    Get-MimeParts $raw $parts
    $files = @(); $plain = $null; $html = $null
    foreach ($p in $parts) {
        if ($p.FileName) {
            if ($p.FileName -ieq 'winmail.dat') { Say '  The email has a winmail.dat attachment (Outlook rich-text format) - can''t unpack that. Save the attachments to a folder instead.' Yellow; continue }
            if ($p.Inline -and $p.Type -like 'image/*') { continue }     # signature logos / pasted pictures
            $out = Get-UniquePath $Dest (Get-SafeFileName $p.FileName)
            [IO.File]::WriteAllBytes($out, $p.Bytes)
            $files += $out
        }
        elseif ($p.Type -eq 'text/plain' -and $null -eq $plain) { $plain = (Get-EncodingSafe $p.Charset).GetString($p.Bytes) }
        elseif ($p.Type -eq 'text/html' -and $null -eq $html) { $html = ConvertFrom-Html ((Get-EncodingSafe $p.Charset).GetString($p.Bytes)) }
    }
    $body = $plain
    if (-not $body) { $body = $html }
    return [pscustomobject]@{ Subject = (ConvertFrom-EncodedWords $top.Headers['subject']); Body = "$body"; Files = $files }
}

function Read-MsgFile([string]$Path, [string]$Dest) {
    $ol = $null; $item = $null
    try {
        try { $ol = New-Object -ComObject Outlook.Application }
        catch { throw "Can't open .msg files without classic Outlook. Use File > Save As (.eml) in Outlook, or save the attachments to a folder and give me the folder." }
        $item = $ol.Session.OpenSharedItem($Path)
        return (Save-OutlookItem $item $Dest)
    }
    finally {
        if ($null -ne $item) { try { $item.Close(1) } catch { } ; Remove-Com $item }
        Remove-Com $ol
    }
}

# Saves an Outlook mail item's attachments (in order, minus embedded pictures)
function Save-OutlookItem($Item, [string]$Dest) {
    $files = @()
    for ($i = 1; $i -le $Item.Attachments.Count; $i++) {
        $a = $Item.Attachments.Item($i)
        $cid = ''
        try { $cid = "$($a.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F'))" } catch { }
        if ($cid -ne '' -and $Script:ImgExt -contains [IO.Path]::GetExtension($a.FileName).ToLower()) { Remove-Com $a; continue }
        $out = Get-UniquePath $Dest (Get-SafeFileName $a.FileName)
        $a.SaveAsFile([string]$out)
        $files += $out
        Remove-Com $a
    }
    return [pscustomobject]@{ Subject = "$($Item.Subject)"; Body = "$($Item.Body)"; Files = $files }
}

# Finds the newest email whose subject contains $Subject (Inbox + its subfolders) in classic Outlook
function Read-OutlookLatest([string]$Subject, [string]$Dest) {
    $ol = $null
    try { $ol = New-Object -ComObject Outlook.Application }
    catch { throw "Couldn't reach Outlook. This needs classic Outlook (Outlook for Windows) set up with your mailbox - new Outlook can't be scripted. Save the email as .eml instead and drag it in." }
    try {
        $ns = $ol.GetNamespace('MAPI')
        $inbox = $ns.GetDefaultFolder(6)        # 6 = Inbox
        # loose match: every main word of the phrase appears somewhere in the subject (any case/spacing),
        # so "Safety Start Up Talks w/o 10/18", "RE: Safety Startup Talks", "safety start-up talk" all count
        $words = @($Subject.ToLower() -split '[^a-z0-9]+' | Where-Object { $_.Length -ge 3 } | ForEach-Object { if ($_.Length -gt 4) { $_.TrimEnd('s') } else { $_ } })
        if ($words.Count -eq 0) { $words = @($Subject.ToLower()) }
        $filter = '@SQL=' + (($words | ForEach-Object { '"urn:schemas:httpmail:subject" LIKE ''%' + $_.Replace("'", "''") + '%''' }) -join ' AND ')
        Say "  Searching Outlook for ""$Subject""..." DarkGray
        $found = New-Object System.Collections.ArrayList
        $stack = New-Object System.Collections.Stack
        $stack.Push($inbox)
        while ($stack.Count -gt 0) {
            $f = $stack.Pop()
            try {
                $items = $f.Items.Restrict($filter)
                foreach ($it in $items) {
                    if ($it.Class -eq 43) { [void]$found.Add([pscustomobject]@{ Item = $it; When = [datetime]$it.ReceivedTime; Folder = "$($f.Name)" }) }   # 43 = mail
                }
            }
            catch { }
            foreach ($sub in $f.Folders) { $stack.Push($sub) }
        }
        if ($found.Count -eq 0) { throw "No email with ""$Subject"" in the subject found in your Inbox (or its folders)." }
        $list = @($found | Sort-Object When -Descending | Select-Object -First 5)
        $pick = $list[0]
        Say ('  Latest: "{0}"  from {1}, {2}  ({3} attachments)' -f $pick.Item.Subject, $pick.Item.SenderName, $pick.When.ToString('ddd MM/dd h:mm tt', $Script:US), $pick.Item.Attachments.Count) White
        if (-not (AskYesNo 'Use this email?' $true)) {
            for ($i = 0; $i -lt $list.Count; $i++) {
                Say ('  [{0}] {1}  -  {2}  ({3})' -f ($i + 1), $list[$i].When.ToString('MM/dd h:mm tt', $Script:US), $list[$i].Item.Subject, $list[$i].Item.SenderName)
            }
            $n = AskInt 'Which one?' 1 $list.Count 1
            $pick = $list[$n - 1]
        }
        return (Save-OutlookItem $pick.Item $Dest)
    }
    finally { Remove-Com $ol }
}

function Get-RecentDownloads {
    if (-not $env:USERPROFILE) { return @() }
    $dl = Join-Path $env:USERPROFILE 'Downloads'
    if (-not (Test-Path -LiteralPath $dl)) { return @() }
    $cut = (Get-Date).AddDays(-14)
    return @(Get-ChildItem -LiteralPath $dl -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in '.eml', '.msg', '.zip' -and $_.LastWriteTime -gt $cut } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 6)
}

# Keeps the order the files came in (email / zip order; folders A-Z)
function Get-TopicFileList([string[]]$Paths) {
    $out = New-Object System.Collections.ArrayList
    $seen = @{}; $skipped = @()
    foreach ($p in $Paths) {
        $items = @()
        if (Test-Path -LiteralPath $p -PathType Container) { $items = @(Get-ChildItem -LiteralPath $p -File -Recurse -ErrorAction SilentlyContinue | Sort-Object Name) }
        elseif (Test-Path -LiteralPath $p) { $items = @(Get-Item -LiteralPath $p) }
        foreach ($f in $items) {
            if ($f.Name.StartsWith('~$') -or $seen.ContainsKey($f.FullName)) { continue }
            $seen[$f.FullName] = $true
            if ($Script:TopicExt -contains $f.Extension.ToLower()) { [void]$out.Add($f) } else { $skipped += $f.Name }
        }
    }
    if ($skipped.Count -gt 0) { Say "  Can't print these, skipped: $($skipped -join ', ')" Yellow }
    return $out.ToArray()
}

# Extracts a zip in its own entry order (Expand-Archive doesn't keep order)
function Expand-ZipInOrder([string]$Zip, [string]$Dest) {
    New-Item -ItemType Directory -Force -Path $Dest | Out-Null
    $files = @()
    $z = [IO.Compression.ZipFile]::OpenRead($Zip)
    try {
        foreach ($e in $z.Entries) {
            if ($e.Name -eq '' -or $e.FullName -match '(^|/)__MACOSX/') { continue }
            $out = Get-UniquePath $Dest (Get-SafeFileName $e.Name)
            [IO.Compression.ZipFileExtensions]::ExtractToFile($e, $out)
            $files += $out
        }
    }
    finally { $z.Dispose() }
    return $files
}

# Returns @{ Subject; Body; Files = FileInfo[]; Mode = 'files'|'manual'|'none'|'reuse'; Map }
function Get-TopicPackage([string[]]$InputPaths, [string]$Work) {
    while ($true) {
        if (-not $InputPaths -or $InputPaths.Count -eq 0) {
            Say 'Drag the email (.eml / .msg), the attachments .zip, the folder with the attachments,'
            Say 'or the attachment files themselves into this window and press Enter. Or pick:'
            $recent = Get-RecentDownloads
            for ($i = 0; $i -lt $recent.Count; $i++) {
                Say ('  [{0}] {1}   ({2})' -f ($i + 1), $recent[$i].Name, $recent[$i].LastWriteTime.ToString('ddd MM/dd h:mm tt', $Script:US))
            }
            $olSubj = Get-Setting (Get-Config) 'OutlookSubject'
            Say ('  [O] Latest "{0}" email from Outlook   (Enter)' -f $olSubj)
            Say '  [R] Reuse topics from a previous export'
            Say '  [M] Type the topic names (no topic sheets)'
            Say '  [S] Skip - leave the topic row blank'
            $a = Ask 'Email / folder / choice' 'O'
            if ($a -ieq 'O') {
                $dest = Join-Path $Work 'outlook'
                New-Item -ItemType Directory -Force -Path $dest | Out-Null
                try { $r = Read-OutlookLatest $olSubj $dest }
                catch { Say "  $($_.Exception.Message)" Yellow; continue }
                $files = @(Get-TopicFileList $r.Files)
                if ($files.Count -eq 0) { Say '  That email has no topic attachments.' Yellow; continue }
                Say "  Subject: $($r.Subject)" DarkGray
                return [pscustomobject]@{ Mode = 'files'; Files = $files; Subject = $r.Subject; Body = $r.Body }
            }
            $n = 0
            if ([int]::TryParse($a, [ref]$n) -and $n -ge 1 -and $n -le $recent.Count) { $InputPaths = @($recent[$n - 1].FullName) }
            elseif ($a -ieq 'M') { return [pscustomobject]@{ Mode = 'manual'; Files = @(); Subject = ''; Body = '' } }
            elseif ($a -ieq 'S') { return [pscustomobject]@{ Mode = 'none'; Files = @(); Subject = ''; Body = '' } }
            elseif ($a -ieq 'R') { $r = Select-PreviousTopics; if ($null -ne $r) { return $r } ; continue }
            else { $InputPaths = @(Split-PathList $a) }
        }
        $missing = @($InputPaths | Where-Object { -not (Test-Path -LiteralPath $_) })
        if ($missing.Count -gt 0) { Say "  Not found: $($missing -join ', ')" Yellow; $InputPaths = @(); continue }

        $subject = ''; $body = ''; $paths = @(); $hadFolder = $false
        try {
            foreach ($p in $InputPaths) {
                $ext = [IO.Path]::GetExtension($p).ToLower()
                if ($ext -eq '.eml' -or $ext -eq '.msg') {
                    $dest = Join-Path $Work 'attachments'
                    New-Item -ItemType Directory -Force -Path $dest | Out-Null
                    Say "  Reading email $([IO.Path]::GetFileName($p)) ..."
                    if ($ext -eq '.eml') { $r = Read-EmlFile $p $dest } else { $r = Read-MsgFile $p $dest }
                    if ($r.Subject) { $subject = $r.Subject; Say "  Subject: $subject" DarkGray }
                    $body += "`n" + $r.Body
                    $paths += $r.Files
                }
                elseif ($ext -eq '.zip') { $paths += @(Expand-ZipInOrder $p (Get-UniquePath $Work 'zip')) }
                else {
                    if (Test-Path -LiteralPath $p -PathType Container) { $hadFolder = $true }
                    $paths += $p
                }
            }
        }
        catch { Say "  $($_.Exception.Message)" Yellow; $InputPaths = @(); continue }

        $files = @(Get-TopicFileList $paths)
        if ($files.Count -eq 0) { Say '  No topic files found there.' Yellow; $InputPaths = @(); continue }
        if ($hadFolder -and $files.Count -gt 6) {
            # a folder like Downloads may hold old files too - default to the ones saved together with the newest
            $newest = ($files | Measure-Object -Property LastWriteTime -Maximum).Maximum
            Say "  Found $($files.Count) files in the folder:"
            $def = @()
            for ($i = 0; $i -lt $files.Count; $i++) {
                Say ('  {0,3}  {1}   ({2})' -f ($i + 1), $files[$i].Name, $files[$i].LastWriteTime.ToString('MM/dd h:mm tt', $Script:US))
                if (($newest - $files[$i].LastWriteTime).TotalMinutes -le 10) { $def += ($i + 1) }
            }
            $pick = Ask 'Which are this week''s topics, in order? (e.g. 1-6 or 3,1,2)' ($def -join ',')
            $idx = @()
            foreach ($tok in ($pick -split '[,\s]+')) {
                if ($tok -match '^(\d+)-(\d+)$') { $idx += ([int]$Matches[1])..([int]$Matches[2]) }
                elseif ($tok -match '^\d+$') { $idx += [int]$tok }
            }
            $files = @($idx | Where-Object { $_ -ge 1 -and $_ -le $files.Count } | ForEach-Object { $files[$_ - 1] })
            if ($files.Count -eq 0) { $InputPaths = @(); continue }
        }
        foreach ($f in $files) { try { Unblock-File -LiteralPath $f.FullName -ErrorAction SilentlyContinue } catch { } }
        return [pscustomobject]@{ Mode = 'files'; Files = $files; Subject = $subject; Body = $body }
    }
}

function Select-PreviousTopics {
    $root = Get-OutputRoot (Get-Config)
    if (-not (Test-Path -LiteralPath $root)) { Say '  No previous exports found.' Yellow; return $null }
    $prev = @(Get-ChildItem -LiteralPath $root -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'topics.json') } | Sort-Object Name -Descending | Select-Object -First 6)
    if ($prev.Count -eq 0) { Say '  No previous exports with saved topics.' Yellow; return $null }
    for ($i = 0; $i -lt $prev.Count; $i++) { Say ('  [{0}] {1}' -f ($i + 1), $prev[$i].Name) }
    $n = AskInt 'Which one?' 1 $prev.Count 1
    $dir = $prev[$n - 1].FullName
    $saved = @(Get-Content -LiteralPath (Join-Path $dir 'topics.json') -Raw | ConvertFrom-Json)
    $map = @()
    for ($d = 0; $d -lt 6; $d++) {
        $e = $saved | Where-Object { $_.Day -eq $Script:DayAbbr[$d] } | Select-Object -First 1
        $file = $null; $title = ''
        if ($e) {
            $title = "$($e.Title)"
            if ($e.File) { $fp = Join-Path (Join-Path $dir 'Topics') $e.File; if (Test-Path -LiteralPath $fp) { $file = Get-Item -LiteralPath $fp } }
        }
        $map += [pscustomobject]@{ Title = $title; File = $file }
    }
    $files = @()
    foreach ($e in $saved) {
        if (-not $e.File) { continue }
        $hit = @($map | Where-Object { $_.File -and $_.File.Name -eq $e.File } | ForEach-Object { $_.File }) | Select-Object -First 1
        if (-not $hit) { $fp = Join-Path (Join-Path $dir 'Topics') $e.File; if (Test-Path -LiteralPath $fp) { $hit = Get-Item -LiteralPath $fp } }
        if ($hit -and $files -notcontains $hit) { $files += $hit }
    }
    return [pscustomobject]@{ Mode = 'reuse'; Files = $files; Subject = ''; Body = ''; Map = $map }
}

#endregion

#region ---------- topics by day ----------

function Get-TopicTitle($File) {
    $t = [IO.Path]::GetFileNameWithoutExtension($File.Name) -replace '_', ' '
    return ([regex]::Replace($t, '\s+', ' ')).Trim()
}

function Get-NormText([string]$s) { return ("$s".ToLower() -replace '[^a-z0-9]', '') }

# Schedule = attachment order: 1st attachment -> Mon, 2nd -> Tue ... 6th -> Sat.
# (Anything past 6 is printed after Sat as an extra.)
function Get-DefaultTopicMap($Files, [string]$Body) {
    $files = @($Files)
    $map = @()
    for ($d = 0; $d -lt 6; $d++) {
        if ($d -lt $files.Count) { $map += [pscustomobject]@{ Title = (Get-TopicTitle $files[$d]); File = $files[$d] } }
        else { $map += [pscustomobject]@{ Title = ''; File = $null } }
    }
    return $map
}

function Get-DayIndex([string]$Tok) {
    if ($Tok.Length -lt 3) { return -1 }
    $t = $Tok.Substring(0, 3).ToLower()
    for ($i = 0; $i -lt 6; $i++) { if ($Script:DayAbbr[$i].ToLower() -eq $t) { return $i } }
    return -1
}

# $Excluded: hashtable of file paths left out of the printed topic packet
function Edit-TopicMap($Files, $Map, $Excluded) {
    $files = @($Files)
    while ($true) {
        Head 'Safety topics by day'
        for ($d = 0; $d -lt 6; $d++) {
            $t = $Map[$d].Title
            if ($t -eq '') { $t = '(none)' }
            $src = ''
            if ($Map[$d].File) { $src = '  [file ' + ([array]::IndexOf($files, $Map[$d].File) + 1) + ']' }
            Say ('  {0}  {1}{2}' -f $Script:DayAbbr[$d], $t, $src) White
        }
        if ($files.Count -gt 0) {
            Say '  Files:' DarkGray
            for ($i = 0; $i -lt $files.Count; $i++) {
                $days = @(); for ($d = 0; $d -lt 6; $d++) { if ($Map[$d].File -eq $files[$i]) { $days += $Script:DayAbbr[$d] } }
                $where = 'extra - printed after the day topics, not on the sheet'
                if ($days.Count -gt 0) { $where = $days -join ', ' }
                if ($Excluded.ContainsKey($files[$i].FullName)) { $where = 'LEFT OUT - not printed' }
                Say ('   {0}. {1}  ({2})' -f ($i + 1), $files[$i].Name, $where) DarkGray
            }
        }
        Say "  Enter = OK.  Change:  'sat 6' file 6 on Sat | 'sat -' nothing on Sat | 'sat Housekeeping' text only" DarkGray
        Say "                        'swap mon tue' | 'title mon New name' | 'skip 7' leave file 7 out (again = put back)" DarkGray
        $a = Ask 'Topics OK?'
        if ($a -eq '') { return $Map }
        $m = [regex]::Match($a, '^(?i)skip\s+(\d+)$')
        if ($m.Success) {
            $n = [int]$m.Groups[1].Value
            if ($n -ge 1 -and $n -le $files.Count) {
                $k = $files[$n - 1].FullName
                if ($Excluded.ContainsKey($k)) { $Excluded.Remove($k) } else { $Excluded[$k] = $true }
            }
            else { Say '  No such file number.' Yellow }
            continue
        }
        $m = [regex]::Match($a, '^(?i)swap\s+(\w+)\s+(\w+)$')
        if ($m.Success) {
            $x = Get-DayIndex $m.Groups[1].Value; $y = Get-DayIndex $m.Groups[2].Value
            if ($x -ge 0 -and $y -ge 0) { $tmp = $Map[$x]; $Map[$x] = $Map[$y]; $Map[$y] = $tmp } else { Say '  Unknown day.' Yellow }
            continue
        }
        $m = [regex]::Match($a, '^(?i)title\s+(\w+)\s+(.+)$')
        if ($m.Success) {
            $x = Get-DayIndex $m.Groups[1].Value
            if ($x -ge 0) { $Map[$x] = [pscustomobject]@{ Title = $m.Groups[2].Value.Trim(); File = $Map[$x].File } } else { Say '  Unknown day.' Yellow }
            continue
        }
        $m = [regex]::Match($a, '^(\w+)\s+(.+)$')
        if ($m.Success) {
            $x = Get-DayIndex $m.Groups[1].Value
            if ($x -lt 0) { Say '  Unknown day.' Yellow; continue }
            $v = $m.Groups[2].Value.Trim()
            $n = 0
            if ($v -eq '-') { $Map[$x] = [pscustomobject]@{ Title = ''; File = $null } }
            elseif ([int]::TryParse($v, [ref]$n) -and $n -ge 1 -and $n -le $files.Count) { $Map[$x] = [pscustomobject]@{ Title = (Get-TopicTitle $files[$n - 1]); File = $files[$n - 1] }; $Excluded.Remove($files[$n - 1].FullName) }
            else { $Map[$x] = [pscustomobject]@{ Title = $v; File = $null } }
            continue
        }
        Say '  Didn''t understand that.' Yellow
    }
}

function Read-ManualTopics {
    $map = @()
    Head 'Type the safety topic for each day (Enter = none)'
    for ($d = 0; $d -lt 6; $d++) { $map += [pscustomobject]@{ Title = (Ask ('  ' + $Script:DayNames[$d])); File = $null } }
    return $map
}

#endregion

#region ---------- Office / PDF ----------

function Import-PdfSharp {
    if ($Script:PdfSharpOK) { return $true }
    if (-not (Test-Path -LiteralPath $Script:PdfSharpPath)) { Say "  Missing $Script:PdfSharpPath" Yellow; return $false }
    try {
        # load from bytes: avoids "downloaded file" blocks and network-drive load rules
        [void][Reflection.Assembly]::Load([IO.File]::ReadAllBytes($Script:PdfSharpPath))
        $Script:PdfSharpOK = $true
        return $true
    }
    catch { Say "  Couldn't load PdfSharp.dll: $($_.Exception.Message)" Yellow; return $false }
}

function Merge-Pdfs([string[]]$Inputs, [string]$Out) {
    $doc = New-Object PdfSharp.Pdf.PdfDocument
    $cache = @{}
    foreach ($p in $Inputs) {
        if (-not $cache.ContainsKey($p)) { $cache[$p] = [PdfSharp.Pdf.IO.PdfReader]::Open($p, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import) }
        $src = $cache[$p]
        for ($i = 0; $i -lt $src.PageCount; $i++) { [void]$doc.AddPage($src.Pages[$i]) }
    }
    $doc.Save($Out)
    $n = $doc.PageCount
    $doc.Dispose()
    foreach ($s in $cache.Values) { $s.Dispose() }
    return $n
}

function Test-PdfReadable([string]$Path) {
    try {
        $d = [PdfSharp.Pdf.IO.PdfReader]::Open($Path, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
        $d.Dispose(); return $true
    }
    catch { return $false }
}

# Topic files -> PDF (Word, PowerPoint, Excel, pictures; PDFs used as-is).
# Returns hashtable: source full path -> pdf path. A file that fails is left out with a warning.
function Convert-TopicsToPdf($Files, [string]$Dir) {
    $map = @{}
    $word = $null; $ppt = $null; $xl = $null
    try {
        foreach ($f in $Files) {
            if ($map.ContainsKey($f.FullName)) { continue }
            $ext = $f.Extension.ToLower()
            if ($ext -eq '.pdf') { $map[$f.FullName] = $f.FullName; continue }
            $pdf = Get-UniquePath $Dir ([IO.Path]::GetFileNameWithoutExtension($f.Name) + '.pdf')
            Say "  Converting $($f.Name)" DarkGray
            try {
                if ($Script:PptExt -contains $ext) {
                    if ($null -eq $ppt) { $ppt = New-Object -ComObject PowerPoint.Application }
                    # Open(FileName, ReadOnly, Untitled, WithWindow)  -1 = true, 0 = false
                    $pres = $ppt.Presentations.Open([string]$f.FullName, -1, 0, 0)
                    try { $pres.SaveAs([string]$pdf, 32) } finally { $pres.Close(); Remove-Com $pres }      # 32 = PDF
                }
                elseif ($Script:XlExt -contains $ext) {
                    if ($null -eq $xl) { $xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false }
                    $wbk = $xl.Workbooks.Open([string]$f.FullName, 0, $true)
                    try { $wbk.ExportAsFixedFormat(0, [string]$pdf) } finally { $wbk.Close($false); Remove-Com $wbk }
                }
                else {
                    if ($null -eq $word) { $word = New-Object -ComObject Word.Application; try { $word.Visible = $false } catch { } }   # (Word's DisplayAlerts is an enum - left alone)
                    if ($Script:ImgExt -contains $ext) {
                        $doc = $word.Documents.Add()
                        $sec = $doc.PageSetup
                        $pic = $doc.InlineShapes.AddPicture([string]$f.FullName)
                        $maxW = $sec.PageWidth - $sec.LeftMargin - $sec.RightMargin
                        $maxH = $sec.PageHeight - $sec.TopMargin - $sec.BottomMargin - 20
                        $scale = [math]::Min(1.0, [math]::Min($maxW / $pic.Width, $maxH / $pic.Height))
                        $pic.LockAspectRatio = -1
                        if ($scale -lt 1.0) { $pic.ScaleWidth = [single](100 * $scale) }
                    }
                    else { $doc = $word.Documents.Open([string]$f.FullName, $false, $true, $false) }
                    try { $doc.ExportAsFixedFormat([string]$pdf, 17) } finally { $doc.Close(0); Remove-Com $doc }   # 17 = PDF
                }
                $map[$f.FullName] = $pdf
            }
            catch { Say "  Couldn't convert $($f.Name): $($_.Exception.Message)" Yellow }
        }
    }
    finally {
        if ($null -ne $word) { try { $word.Quit(0) } catch { } ; Remove-Com $word }
        if ($null -ne $ppt) { try { $ppt.Quit() } catch { } ; Remove-Com $ppt }
        if ($null -ne $xl) { try { $xl.Quit() } catch { } ; Remove-Com $xl }
    }
    return $map
}

function Set-SignSheet($Ws, $Cfg, $Shop, [int]$Shift, [datetime]$Monday, $Topics) {
    $night = Get-SettingInt $Cfg 'NightShift' 3
    $start = $Monday
    if ($Shift -eq $night) { $start = $Monday.AddDays(-1) }
    $emps = @(Get-ShopEmployees $Cfg $Shop.Code $Shift)

    # NOTE: Windows PowerShell 5.1 caches COM property-set rules by name, so every .Value2
    # write here is a [string] (mixing doubles/arrays with strings throws "Unable to cast").
    # The date goes in as a formula instead.
    $Ws.Range('A1').Value2 = [string]$Shop.Title
    $Ws.Range('F1').Value2 = [string](Get-ShiftName $Cfg $Shift)
    $Ws.Range('I1').Formula = [string]('=DATE({0},{1},{2})' -f $start.Year, $start.Month, $start.Day)
    $Ws.Range('C2:H3').NumberFormat = '@'
    for ($d = 0; $d -lt 6; $d++) {
        $Ws.Cells.Item(2, 3 + $d).Value2 = [string]($Script:DayAbbr[$d] + '  ' + $start.AddDays($d).ToString('MM/dd/yy', $Script:Inv))
        $Ws.Cells.Item(3, 3 + $d).Value2 = [string]"$($Topics[$d].Title)"
    }

    $need = [math]::Max($emps.Count + (Get-SettingInt $Cfg 'ExtraBlankRows' 3), (Get-SettingInt $Cfg 'MinRows' 10))
    if ($need -lt 3) { $need = 3 }
    $tplRows = $Ws.Range('I5').MergeArea.Rows.Count
    if ($need -lt $tplRows) { [void]$Ws.Range('A6:A' + (5 + $tplRows - $need)).EntireRow.Delete() }
    elseif ($need -gt $tplRows) { [void]$Ws.Range('A7:A' + (6 + $need - $tplRows)).EntireRow.Insert() }
    $last = 4 + $need
    [void]$Ws.Range("A5:B$last").ClearContents()
    $Ws.Range("A5:B$last").NumberFormat = '@'
    for ($i = 0; $i -lt $emps.Count; $i++) {
        $Ws.Cells.Item(5 + $i, 1).Value2 = [string]$emps[$i].Name
        $Ws.Cells.Item(5 + $i, 2).Value2 = [string]$emps[$i].Assignment
    }
    $Ws.PageSetup.PrintArea = [string]('$A$1:$I$' + $last)
    return $emps.Count
}

# Builds every shop/shift sheet from the template, then exports.
#   -OutPdf      : one PDF with everything
#   -PerShopDir  : one PDF per shop in that folder; returns ordered dict code -> path
function Export-SignSheets($Cfg, $Shops, [datetime]$Monday, $Topics, [string]$OutPdf, [string]$PerShopDir) {
    if (-not (Test-Path -LiteralPath $Script:TemplatePath)) { throw "Template not found: $Script:TemplatePath" }
    $xl = $null; $wb = $null
    $result = [ordered]@{}
    try {
        Say '  Starting Excel...' DarkGray
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.ScreenUpdating = $false
        $wb = $xl.Workbooks.Open($Script:TemplatePath, 0, $true)
        $tpl = $wb.Worksheets.Item(1)
        if ($tpl.Range('I5').MergeArea.Rows.Count -lt 3) { throw 'Template problem: I5 must be merged down across the employee rows (3+ rows).' }
        $sheetsByShop = [ordered]@{}
        foreach ($shop in $Shops) {
            $names = @()
            for ($s = 1; $s -le $shop.Shifts; $s++) {
                $tpl.Copy([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
                $ws = $wb.Worksheets.Item($wb.Worksheets.Count)
                $nm = ('{0} {1}' -f $shop.Code, (Get-Ordinal $s))
                if ($nm.Length -gt 31) { $nm = $nm.Substring(0, 31) }
                $ws.Name = [string]$nm
                $count = Set-SignSheet $ws $Cfg $shop $s $Monday $Topics
                Say ('  {0,-6} {1,-13} {2} people' -f $shop.Code, (Get-ShiftName $Cfg $s), $count) DarkGray
                $names += $ws.Name
                Remove-Com $ws
            }
            $sheetsByShop[$shop.Code] = $names
        }
        $tpl.Delete()
        Remove-Com $tpl
        if ($PerShopDir) {
            foreach ($code in $sheetsByShop.Keys) {
                foreach ($n in $sheetsByShop[$code]) { $wb.Worksheets.Item($n).Visible = $true }
                foreach ($ws in $wb.Worksheets) { if ($sheetsByShop[$code] -notcontains $ws.Name) { $ws.Visible = $false } }
                $p = Join-Path $PerShopDir ("signsheets_$code.pdf")
                $wb.ExportAsFixedFormat(0, $p, 0, $false, $false)
                $result[$code] = $p
            }
        }
        else {
            $wb.ExportAsFixedFormat(0, $OutPdf, 0, $false, $false)
        }
    }
    finally {
        if ($null -ne $wb) { try { $wb.Close($false) } catch { } ; Remove-Com $wb }
        if ($null -ne $xl) { try { $xl.Quit() } catch { } ; Remove-Com $xl }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    return $result
}

#endregion

#region ---------- export ----------

function Select-Shops($Cfg) {
    Say '  [A] All shops'
    for ($i = 0; $i -lt $Cfg.Shops.Count; $i++) {
        $s = $Cfg.Shops[$i]
        $n = @($Cfg.Employees | Where-Object { $_.Shop -eq $s.Code -and $_.Shift -le $s.Shifts }).Count
        Say ('  [{0}] {1,-5} {2}  ({3} shift{4}, {5} people)' -f ($i + 1), $s.Code, $s.Title, $s.Shifts, $(if ($s.Shifts -ne 1) { 's' } else { '' }), $n)
    }
    while ($true) {
        $a = Ask 'Which shops? (A, or e.g. 1  /  HMS  /  1,3)' 'A'
        if ($a -ieq 'A' -or $a -ieq 'all') { return @($Cfg.Shops) }
        $sel = @(); $bad = $false
        foreach ($tok in ($a -split '[,\s]+' | Where-Object { $_ -ne '' })) {
            $s = Find-Shop $Cfg $tok
            if ($null -eq $s) { Say "  Unknown shop '$tok'." Yellow; $bad = $true; break }
            if ($sel -notcontains $s) { $sel += $s }
        }
        if (-not $bad -and $sel.Count -gt 0) { return $sel }
    }
}

function Invoke-Export([string[]]$InputPaths) {
    $cfg = Get-Config
    if ($cfg.Shops.Count -eq 0) { throw 'No shops in HuddleConfig.xlsx - add one with add-shop.' }
    $work = Join-Path ([IO.Path]::GetTempPath()) ('HuddleSheets_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    try {
        Head 'STEP 1 of 4 - Safety topics'
        $pkg = Get-TopicPackage $InputPaths $work
        $excluded = @{}
        if ($pkg.Mode -eq 'files') { $map = Edit-TopicMap $pkg.Files (Get-DefaultTopicMap $pkg.Files $pkg.Body) $excluded }
        elseif ($pkg.Mode -eq 'reuse') { $map = Edit-TopicMap $pkg.Files $pkg.Map $excluded }
        elseif ($pkg.Mode -eq 'manual') { $map = Read-ManualTopics }
        else { $map = @(); for ($d = 0; $d -lt 6; $d++) { $map += [pscustomobject]@{ Title = ''; File = $null } } }

        Head 'STEP 2 of 4 - Week'
        $monday = Select-Week (Get-WeekFromSubject $pkg.Subject)
        $weekNo = Get-WeekNumber $monday

        Head 'STEP 3 of 4 - Shops'
        $shops = @(Select-Shops $cfg)
        $all = ($shops.Count -eq $cfg.Shops.Count)

        $withTopics = $false; $copies = 0; $afterEach = $true
        # printed topic packet: day topics Mon-Sat, then any extra files, in the order received
        $packetFiles = @()
        foreach ($t in $map) { if ($t.File -and $packetFiles -notcontains $t.File) { $packetFiles += $t.File } }
        foreach ($f in @($pkg.Files)) { if ($packetFiles -notcontains $f -and -not $excluded.ContainsKey($f.FullName)) { $packetFiles += $f } }
        Head 'STEP 4 of 4 - Safety topic sheets'
        if ($packetFiles.Count -gt 0) {
            Say '  Topic sheets print in schedule order:' DarkGray
            for ($i = 0; $i -lt $packetFiles.Count; $i++) {
                $day = 'extra'
                for ($d = 0; $d -lt 6; $d++) { if ($map[$d].File -eq $packetFiles[$i]) { $day = $Script:DayAbbr[$d]; break } }
                Say ('    {0}. {1,-5} {2}' -f ($i + 1), $day, $packetFiles[$i].Name) DarkGray
            }
            $withTopics = AskYesNo 'Include the safety topic sheets in the PDF?' $false
            if ($withTopics) {
                if ($shops.Count -gt 1) {
                    Say '  [1] After each shop''s sign-in sheets   [2] Once, at the end'
                    $afterEach = ((AskInt 'Where?' 1 2 1) -eq 1)
                }
                $q = 'How many copies of the topic sheets?'
                if ($afterEach -and $shops.Count -gt 1) { $q = 'How many copies of the topic sheets per shop?' }
                $copies = AskInt $q 1 50 (Get-SettingInt $cfg 'TopicCopies' 1)
            }
        }
        else { Say '  No topic documents - sign-in sheets only.' DarkGray }

        # output location
        $weekDir = Join-Path (Get-OutputRoot $cfg) ('Week {0:00} - {1}' -f $weekNo, $monday.ToString('yyyy-MM-dd', $Script:Inv))
        New-Item -ItemType Directory -Force -Path $weekDir | Out-Null
        $label = ($shops | ForEach-Object { $_.Code }) -join ', '
        if ($all) { $label = 'All Shops' }
        $fileName = 'Huddle Sheets - Week {0} - {1}{2}.pdf' -f $weekNo, $label, $(if ($withTopics) { ' + Topics' } else { '' })
        $final = Join-Path $weekDir $fileName
        if (Test-Path -LiteralPath $final) {
            try { Remove-Item -LiteralPath $final -Force } catch { $final = Get-UniquePath $weekDir $fileName; Say '  Old PDF is open somewhere - saving under a new name.' Yellow }
        }

        Head 'Building...'
        $pages = 0
        if ($withTopics -and -not (Import-PdfSharp)) {
            Say '  Can''t combine PDFs - making the sign-in sheets only.' Yellow
            $withTopics = $false
        }
        if (-not $withTopics) {
            [void](Export-SignSheets $cfg $shops $monday $map $final $null)
        }
        else {
            $shopPdfs = Export-SignSheets $cfg $shops $monday $map $null $work
            $conv = Convert-TopicsToPdf $packetFiles $work
            $packet = @()
            foreach ($f in $packetFiles) {
                $p = $conv[$f.FullName]
                if ($p -and (Test-PdfReadable $p)) { $packet += $p }
                else { Say "  Skipping $($f.Name) - couldn't convert / read it." Yellow }
            }
            $list = @()
            foreach ($code in $shopPdfs.Keys) {
                $list += $shopPdfs[$code]
                if ($afterEach) { for ($c = 0; $c -lt $copies; $c++) { $list += $packet } }
            }
            if (-not $afterEach) { for ($c = 0; $c -lt $copies; $c++) { $list += $packet } }
            $pages = Merge-Pdfs $list $final
        }

        # keep the topic documents + day mapping with the week (used by "Reuse topics")
        $saved = @()
        $tdir = Join-Path $weekDir 'Topics'
        foreach ($f in @($pkg.Files)) {
            New-Item -ItemType Directory -Force -Path $tdir | Out-Null
            $dest = Join-Path $tdir $f.Name
            if ($f.FullName -ne $dest) { Copy-Item -LiteralPath $f.FullName -Destination $dest -Force }
        }
        for ($d = 0; $d -lt 6; $d++) {
            $fileOut = $null
            if ($map[$d].File) { $fileOut = $map[$d].File.Name }
            $saved += [pscustomobject]@{ Day = $Script:DayAbbr[$d]; Title = "$($map[$d].Title)"; File = $fileOut }
        }
        foreach ($f in $packetFiles) {
            if (@($map | Where-Object { $_.File -eq $f }).Count -eq 0) { $saved += [pscustomobject]@{ Day = 'Extra'; Title = ''; File = $f.Name } }
        }
        $saved | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $weekDir 'topics.json') -Encoding UTF8

        Head 'Done'
        $info = ''
        if ($pages -gt 0) { $info = " ($pages pages)" }
        Say "  $final$info" Green
        if ((Get-Setting $cfg 'OpenWhenDone') -match '^(y|yes|true|1)$') { try { Invoke-Item -LiteralPath $final } catch { } }
    }
    finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}

#endregion

#region ---------- config commands ----------

function Read-Shop($Cfg, [string]$Default = '') {
    if ($Cfg.Shops.Count -eq 1) { return $Cfg.Shops[0] }
    $opts = for ($i = 0; $i -lt $Cfg.Shops.Count; $i++) { '{0}={1}' -f ($i + 1), $Cfg.Shops[$i].Code }
    while ($true) {
        $a = Ask ("Shop (" + ($opts -join ', ') + ')') $Default
        $s = Find-Shop $Cfg $a
        if ($null -ne $s) { return $s }
        Say '  Unknown shop.' Yellow
    }
}

function Read-Shift($Shop, [int]$Default = 1) {
    if ($Shop.Shifts -le 1) { return 1 }
    if ($Default -gt $Shop.Shifts) { $Default = 1 }
    return (AskInt "Shift (1-$($Shop.Shifts))" 1 $Shop.Shifts $Default)
}

function Read-Assignment($Cfg, [string]$Default = '') {
    $list = @($Cfg.Employees | Where-Object { $_.Assignment } | Group-Object Assignment | Sort-Object Count -Descending | ForEach-Object { $_.Name })
    if ($list.Count -gt 0) {
        $line = for ($i = 0; $i -lt $list.Count; $i++) { '{0}) {1}' -f ($i + 1), $list[$i] }
        Say ('  ' + ($line -join '   ')) DarkGray
    }
    if ($Default -eq '' -and $list.Count -gt 0) { $Default = $list[0] }
    $a = Ask 'Assignment (number or type a new one)' $Default
    $n = 0
    if ([int]::TryParse($a, [ref]$n) -and $n -ge 1 -and $n -le $list.Count) { return $list[$n - 1] }
    return $a
}

function Format-Name([string]$Last, [string]$First) {
    $l = ([regex]::Replace($Last.Trim(), '\s+', ' ')).ToUpper()
    $f = ([regex]::Replace($First.Trim(), '\s+', ' ')).ToUpper()
    if ($f -eq '') { return $l }
    return "$l, $f"
}

function Add-EmployeeToConfig($Cfg, $Emp) {
    # supervisors go to the top of their shift, everyone else to the bottom
    $idx = -1
    for ($i = 0; $i -lt $Cfg.Employees.Count; $i++) {
        $e = $Cfg.Employees[$i]
        if ($e.Shop -eq $Emp.Shop -and $e.Shift -eq $Emp.Shift) {
            if ($Emp.Assignment -match 'supervisor') { if ($idx -lt 0) { $idx = $i } }
            else { $idx = $i + 1 }
        }
    }
    if ($idx -lt 0) { [void]$Cfg.Employees.Add($Emp) } else { $Cfg.Employees.Insert($idx, $Emp) }
}

function Select-Employee($Cfg, [string]$Query) {
    if ($Query -eq '') { $Query = Ask 'Name (or part of it)' }
    if ($Query -eq '') { throw (New-Object OperationCanceledException) }
    $hits = @($Cfg.Employees | Where-Object { $_.Name -like "*$Query*" })
    if ($hits.Count -eq 0) { Say "  Nobody matches '$Query'." Yellow; return $null }
    if ($hits.Count -eq 1) { return $hits[0] }
    for ($i = 0; $i -lt $hits.Count; $i++) { Say ('  [{0}] {1,-28} {2,-5} {3,-13} {4}' -f ($i + 1), $hits[$i].Name, $hits[$i].Shop, (Get-ShiftName $Cfg $hits[$i].Shift), $hits[$i].Assignment) }
    $n = AskInt 'Which one?' 1 $hits.Count 1
    return $hits[$n - 1]
}

function Show-Employees($Cfg, [string[]]$ArgList) {
    $shops = @($Cfg.Shops)
    $shiftFilter = 0
    foreach ($a in $ArgList) {
        $n = 0
        if ([int]::TryParse($a, [ref]$n) -and $n -ge 1 -and $n -le 3 -and $null -eq ($Cfg.Shops | Where-Object { $_.Code -eq $a })) { $shiftFilter = $n; continue }
        $s = Find-Shop $Cfg $a
        if ($null -eq $s) { Say "  Unknown shop '$a'." Yellow; return }
        $shops = @($s)
    }
    foreach ($s in $shops) {
        Head ("{0} - {1}" -f $s.Code, $s.Title)
        for ($sh = 1; $sh -le $s.Shifts; $sh++) {
            if ($shiftFilter -gt 0 -and $sh -ne $shiftFilter) { continue }
            $emps = @(Get-ShopEmployees $Cfg $s.Code $sh)
            Say ("  {0} ({1})" -f (Get-ShiftName $Cfg $sh), $emps.Count) White
            for ($i = 0; $i -lt $emps.Count; $i++) { Say ('    {0,2}. {1,-30} {2}' -f ($i + 1), $emps[$i].Name, $emps[$i].Assignment) }
        }
        $orphans = @($Cfg.Employees | Where-Object { $_.Shop -eq $s.Code -and $_.Shift -gt $s.Shifts })
        if ($orphans.Count -gt 0) { Say "  ! $($orphans.Count) people are on a shift above $($s.Shifts) and won't print: $(($orphans | ForEach-Object { $_.Name }) -join '; ')" Yellow }
    }
    $lost = @($Cfg.Employees | Where-Object { $e = $_; -not ($Cfg.Shops | Where-Object { $_.Code -eq $e.Shop }) })
    if ($lost.Count -gt 0) { Say "  ! $($lost.Count) people have a shop code that isn't on the Shops sheet: $(($lost | ForEach-Object { "$($_.Name) ($($_.Shop))" }) -join '; ')" Yellow }
}

function Invoke-AddEmployee {
    $cfg = Get-Config
    Head 'Add employee'
    $shop = Read-Shop $cfg
    $shift = Read-Shift $shop 1
    $last = ''
    while ($last -eq '') { $last = Ask 'Last name' }
    $first = Ask 'First name'
    $name = Format-Name $last $first
    $asg = Read-Assignment $cfg
    $dup = @($cfg.Employees | Where-Object { $_.Shop -eq $shop.Code -and $_.Shift -eq $shift -and $_.Name -ieq $name })
    if ($dup.Count -gt 0) { Say "  $name is already on $($shop.Code) $(Get-ShiftName $cfg $shift)." Yellow; return }
    $other = @($cfg.Employees | Where-Object { $_.Name -ieq $name })
    foreach ($o in $other) { Say "  Note: $name is also on $($o.Shop) $(Get-ShiftName $cfg $o.Shift)." DarkGray }
    Add-EmployeeToConfig $cfg (New-Employee $shop.Code $shift $name $asg)
    Save-Config $cfg
    Say "  Added $name - $asg - $($shop.Code) $(Get-ShiftName $cfg $shift)" Green
}

function Invoke-RemoveEmployee([string]$Query) {
    $cfg = Get-Config
    $e = Select-Employee $cfg $Query
    if ($null -eq $e) { return }
    if (AskYesNo "Remove $($e.Name) from $($e.Shop) $(Get-ShiftName $cfg $e.Shift)?" $true) {
        $cfg.Employees.Remove($e)
        Save-Config $cfg
        Say "  Removed $($e.Name)." Green
    }
}

function Invoke-EditEmployee([string]$Query) {
    $cfg = Get-Config
    $e = Select-Employee $cfg $Query
    if ($null -eq $e) { return }
    Head "Edit $($e.Name)  (Enter keeps the current value)"
    $name = (Ask 'Name (LAST, FIRST)' $e.Name).ToUpper()
    $asg = Read-Assignment $cfg $e.Assignment
    $shop = Read-Shop $cfg $e.Shop
    $shift = Read-Shift $shop $e.Shift
    $moved = ($shop.Code -ne $e.Shop -or $shift -ne $e.Shift)
    $pos = $cfg.Employees.IndexOf($e)
    $cfg.Employees.Remove($e)
    $new = New-Employee $shop.Code $shift $name $asg
    if ($moved) { Add-EmployeeToConfig $cfg $new }
    else { $cfg.Employees.Insert($pos, $new) }      # keep their place in the list
    Save-Config $cfg
    Say "  Saved: $name - $asg - $($shop.Code) $(Get-ShiftName $cfg $shift)" Green
}

function Show-Shops($Cfg) {
    Head 'Shops'
    for ($i = 0; $i -lt $Cfg.Shops.Count; $i++) {
        $s = $Cfg.Shops[$i]
        $n = @($Cfg.Employees | Where-Object { $_.Shop -eq $s.Code }).Count
        Say ('  {0}. {1,-5} {2,-50} {3} shift(s), {4} people' -f ($i + 1), $s.Code, $s.Title, $s.Shifts, $n)
    }
}

function Invoke-AddShop {
    $cfg = Get-Config
    Head 'Add shop'
    $code = (Ask 'Shop code (short, e.g. FAB)').ToUpper() -replace '\s', ''
    if ($code -eq '') { return }
    if ($null -ne (Find-Shop $cfg $code)) { Say "  $code already exists." Yellow; return }
    $title = Ask 'Title printed on the sheet' "$code SST & Shift Huddle"
    $shifts = AskInt 'How many shifts?' 1 3 2
    [void]$cfg.Shops.Add([pscustomobject]@{ Code = $code; Title = $title; Shifts = $shifts })
    Save-Config $cfg
    Say "  Added $code. Add people with 'add'." Green
}

function Invoke-RemoveShop([string]$Code) {
    $cfg = Get-Config
    $s = Find-Shop $cfg $Code
    if ($null -eq $s) { $s = Read-Shop $cfg }
    $emps = @($cfg.Employees | Where-Object { $_.Shop -eq $s.Code })
    if (-not (AskYesNo "Remove shop $($s.Code) and its $($emps.Count) people?" $false)) { return }
    $cfg.Shops.Remove($s)
    foreach ($e in $emps) { $cfg.Employees.Remove($e) }
    Save-Config $cfg
    Say "  Removed $($s.Code)." Green
}

function Invoke-SetShifts([string[]]$ArgList) {
    $cfg = Get-Config
    $s = $null; $n = 0
    if ($ArgList.Count -ge 1) { $s = Find-Shop $cfg $ArgList[0] }
    if ($null -eq $s) { $s = Read-Shop $cfg }
    if ($ArgList.Count -ge 2 -and [int]::TryParse($ArgList[1], [ref]$n) -and $n -ge 1 -and $n -le 3) { }
    else { $n = AskInt "How many shifts for $($s.Code)?" 1 3 $s.Shifts }
    $orphans = @($cfg.Employees | Where-Object { $_.Shop -eq $s.Code -and $_.Shift -gt $n })
    if ($orphans.Count -gt 0) {
        Say "  $($orphans.Count) people are on shifts above $n." Yellow
        Say '  [1] Move them to shift 1   [2] Delete them   [3] Leave them (they won''t print)'
        $c = AskInt 'Choose' 1 3 1
        foreach ($e in $orphans) {
            if ($c -eq 1) { $cfg.Employees.Remove($e); $e.Shift = 1; Add-EmployeeToConfig $cfg $e }
            elseif ($c -eq 2) { $cfg.Employees.Remove($e) }
        }
    }
    $s.Shifts = $n
    Save-Config $cfg
    Say "  $($s.Code) now has $n shift(s)." Green
}

function Invoke-SetTitle([string]$Code) {
    $cfg = Get-Config
    $s = Find-Shop $cfg $Code
    if ($null -eq $s) { $s = Read-Shop $cfg }
    $s.Title = Ask 'Title printed on the sheet' $s.Title
    Save-Config $cfg
    Say '  Saved.' Green
}

function Show-Settings($Cfg) {
    Head 'Settings  (change with: set <name> <value>)'
    foreach ($k in $Cfg.Settings.Keys) {
        $v = $Cfg.Settings[$k]; if ($v -eq '') { $v = '(blank)' }
        Say ('  {0,-15} {1}' -f $k, $v) White
        if ($Script:DefaultSettings.Contains($k)) { Say ('  {0,-15} {1}' -f '', $Script:DefaultSettings[$k][1]) DarkGray }
    }
    Say ('  Output folder now: ' + (Get-OutputRoot $Cfg)) DarkGray
}

function Invoke-Set([string[]]$ArgList) {
    $cfg = Get-Config
    if ($ArgList.Count -lt 1) { Show-Settings $cfg; return }
    $key = $null
    foreach ($k in $cfg.Settings.Keys) { if ($k -ieq $ArgList[0]) { $key = $k } }
    if ($null -eq $key) { Say "  Unknown setting '$($ArgList[0])'. Type 'settings' to see them." Yellow; return }
    if ($ArgList.Count -ge 2) { $val = ($ArgList[1..($ArgList.Count - 1)] -join ' ') } else { $val = Ask "New value for $key" $cfg.Settings[$key] }
    if ($val -eq '""' -or $val -eq "''") { $val = '' }
    $cfg.Settings[$key] = $val
    Save-Config $cfg
    Say "  $key = $val" Green
}

function Show-Weeks {
    Head 'Weeks'
    $this = Get-WeekMonday (Get-Date)
    for ($i = -3; $i -le 4; $i++) {
        $m = $this.AddDays(7 * $i)
        $mark = ''; if ($i -eq 0) { $mark = '  <- this week' }
        Say ('  {0}   (3rd shift starts Sun {1}){2}' -f (Format-Week $m), $m.AddDays(-1).ToString('MM/dd', $Script:Inv), $mark)
    }
}

function Show-Help {
    Head 'Commands  (Enter on an empty line = export)'
    $rows = @(
        @('export [email/zip/folder]', 'Build the weekly PDF: topics -> week -> shops -> topic copies', 'e'),
        @('weeks', 'Week numbers around today', ''),
        @('list [shop] [shift]', "List people, e.g. 'list', 'list hms', 'list hms 3'", 'ls'),
        @('find <text>', 'Search people by name', ''),
        @('add', 'Add an employee (asks shop, shift, name, assignment)', ''),
        @('remove [name]', 'Remove an employee', 'rm'),
        @('edit [name]', 'Change name, assignment, shop or shift', 'move'),
        @('shops', 'List shops and shift counts', ''),
        @('add-shop', 'Add a shop', ''),
        @('remove-shop [shop]', 'Remove a shop and its people', ''),
        @('shifts [shop] [n]', "Set number of shifts, e.g. 'shifts mms 3'", ''),
        @('title [shop]', 'Change the title printed on a shop''s sheets', ''),
        @('settings', 'Show settings', ''),
        @('set <name> <value>', "Change a setting, e.g. 'set ExtraBlankRows 4'", ''),
        @('open config|template|output', 'Open the config workbook, the sheet template, or the PDF folder', ''),
        @('help', 'This list', '?'),
        @('exit', 'Quit', 'q')
    )
    foreach ($r in $rows) {
        $alias = ''; if ($r[2]) { $alias = "  [$($r[2])]" }
        Say ('  {0,-28} {1}{2}' -f $r[0], $r[1], $alias)
    }
    Say "  Type 'cancel' at any prompt to back out. People/shops can also be edited directly in HuddleConfig.xlsx." DarkGray
}

function Invoke-HuddleCommand([string]$Line) {
    $tokens = @(Split-PathList $Line)
    if ($tokens.Count -eq 0) { return }
    $cmd = $tokens[0].ToLower()
    $rest = @()
    if ($tokens.Count -gt 1) { $rest = @($tokens[1..($tokens.Count - 1)]) }
    $restText = ($rest -join ' ')
    $alias = @{ e = 'export'; x = 'export'; run = 'export'; ls = 'list'; employees = 'list'; rm = 'remove'; del = 'remove'; delete = 'remove'
        move = 'edit'; mv = 'edit'; '?' = 'help'; h = 'help'; q = 'exit'; quit = 'exit'; setting = 'settings'; config = 'settings'
        week = 'weeks'; search = 'find'; 'add-employee' = 'add'; 'remove-employee' = 'remove'; 'edit-employee' = 'edit'; 'set-shifts' = 'shifts'; 'rename-shop' = 'title'
    }
    if ($alias.ContainsKey($cmd)) { $cmd = $alias[$cmd] }

    # a dragged-in path (or several) = export from it
    $whole = Clean-PathText $Line
    if ($cmd -ne 'export' -and [IO.Path]::IsPathRooted($tokens[0]) -and (Test-Path -LiteralPath $tokens[0] -ErrorAction SilentlyContinue)) { Invoke-Export $tokens; return }

    switch ($cmd) {
        'export' {
            $p = Clean-PathText $restText
            if ($p -ne '' -and (Test-Path -LiteralPath $p -ErrorAction SilentlyContinue)) { Invoke-Export @($p) }
            elseif ($rest.Count -gt 0) { Invoke-Export $rest }
            else { Invoke-Export @() }
        }
        'list' { Show-Employees (Get-Config) $rest }
        'find' {
            $cfg = Get-Config
            $q = $restText; if ($q -eq '') { $q = Ask 'Name (or part of it)' }
            $hits = @($cfg.Employees | Where-Object { $_.Name -like "*$q*" })
            if ($hits.Count -eq 0) { Say '  No match.' Yellow }
            foreach ($e in $hits) { Say ('  {0,-30} {1,-5} {2,-13} {3}' -f $e.Name, $e.Shop, (Get-ShiftName $cfg $e.Shift), $e.Assignment) }
        }
        'add' { Invoke-AddEmployee }
        'remove' { Invoke-RemoveEmployee $restText }
        'edit' { Invoke-EditEmployee $restText }
        'shops' { Show-Shops (Get-Config) }
        'add-shop' { Invoke-AddShop }
        'remove-shop' { Invoke-RemoveShop $restText }
        'shifts' { Invoke-SetShifts $rest }
        'title' { Invoke-SetTitle $restText }
        'settings' { Show-Settings (Get-Config) }
        'set' { Invoke-Set $rest }
        'weeks' { Show-Weeks }
        'open' {
            $what = "$restText".ToLower()
            if ($what -like 'temp*') { Invoke-Item -LiteralPath $Script:TemplatePath }
            elseif ($what -like 'out*' -or $what -like 'pdf*') { $o = Get-OutputRoot (Get-Config); New-Item -ItemType Directory -Force -Path $o | Out-Null; Invoke-Item -LiteralPath $o }
            else { Invoke-Item -LiteralPath $Script:ConfigPath }
        }
        'help' { Show-Help }
        'exit' { return 'exit' }
        default { Say "  Unknown command '$cmd'. Type 'help'." Yellow }
    }
}

function Start-Repl {
    Say ''
    Say '  SST & SHIFT HUDDLE SHEETS' Cyan
    Say '  Press Enter to build this week''s PDF, or type a command (help = list of commands).' DarkGray
    while ($true) {
        Write-Host ''
        $line = Read-Host 'huddle'
        if ($null -eq $line) { break }
        $line = $line.Trim()
        if ($line -eq '') { $line = 'export' }
        try { if ((Invoke-HuddleCommand $line) -eq 'exit') { break } }
        catch [OperationCanceledException] { Say '  Cancelled.' Yellow }
        catch { Show-Error $_ }
    }
}

#endregion

# ---------- entry point ----------
if ($MyInvocation.InvocationName -ne '.') {
    if ($CmdArgs -and $CmdArgs.Count -gt 0) {
        $paths = @($CmdArgs | Where-Object { Test-Path -LiteralPath $_ -ErrorAction SilentlyContinue })
        try {
            if ($paths.Count -eq $CmdArgs.Count) { Invoke-Export $paths }      # files dragged onto the .bat
            else { [void](Invoke-HuddleCommand ($CmdArgs -join ' ')) }
        }
        catch [OperationCanceledException] { Say '  Cancelled.' Yellow }
        catch { Show-Error $_ }
        if ($paths.Count -eq $CmdArgs.Count) { Start-Repl }
    }
    else { Start-Repl }
}
