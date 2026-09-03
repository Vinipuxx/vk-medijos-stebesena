<#
.SYNOPSIS
  Sukuria .xlsx failą iš JSON aprašo (be Python/Node priklausomybių, tik .NET).

.PARAMETER InputJson
  Kelias iki JSON failo su struktūra:
  {
    "sheetName": "Apzvalga",
    "columns": [ { "header": "Nr.", "key": "nr", "width": 6, "wrap": false, "hyperlink": false, "numeric": false } ... ],
    "rows": [ { "nr": 1, "pavadinimas": "..." }, ... ]
  }
  Jei stulpelio "hyperlink" = true, atitinkamo rakto reikšmė laikoma URL ir langelis tampa spustelėjama nuoroda
  (rodomas tekstas imamas iš "<key>_text", jei toks yra, kitaip rodomas pats URL).

.PARAMETER OutputXlsx
  Kelias, kur išsaugoti sukurtą .xlsx failą.
#>
param(
    [Parameter(Mandatory=$true)][string]$InputJson,
    [Parameter(Mandatory=$true)][string]$OutputXlsx
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Esc([string]$s) {
    if ($null -eq $s) { return "" }
    $s = $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
    # strip control chars invalid in XML 1.0
    $s = -join ($s.ToCharArray() | Where-Object { ([int]$_) -ge 32 -or [int]$_ -eq 9 -or [int]$_ -eq 10 -or [int]$_ -eq 13 })
    return $s
}

$spec = Get-Content -Raw -Path $InputJson -Encoding UTF8 | ConvertFrom-Json
$sheetName = if ($spec.sheetName) { $spec.sheetName } else { "Sheet1" }
$columns = $spec.columns
$rows = $spec.rows

# ---- shared strings ----
$sharedStrings = New-Object System.Collections.Generic.List[string]
$stringIndex = @{}
function Get-StringIndex([string]$s) {
    if (-not $stringIndex.ContainsKey($s)) {
        $stringIndex[$s] = $sharedStrings.Count
        $sharedStrings.Add($s) | Out-Null
    }
    return $stringIndex[$s]
}

function ColLetter([int]$n) {
    $letters = ""
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        $letters = [char](65 + $rem) + $letters
        $n = [int](($n - $rem - 1) / 26)
    }
    return $letters
}

# ---- build sheet rows ----
$sb = New-Object System.Text.StringBuilder
$rowIdx = 1

# header row
[void]$sb.Append("<row r=`"$rowIdx`" s=`"2`">")
for ($c = 0; $c -lt $columns.Count; $c++) {
    $col = $columns[$c]
    $ref = (ColLetter ($c+1)) + $rowIdx
    $si = Get-StringIndex (Esc $col.header)
    [void]$sb.Append("<c r=`"$ref`" t=`"s`" s=`"2`"><v>$si</v></c>")
}
[void]$sb.Append("</row>")
$rowIdx++

$hyperlinkEntries = New-Object System.Collections.Generic.List[string]
$hlRelId = 1
$hlRels = New-Object System.Collections.Generic.List[string]

foreach ($row in $rows) {
    [void]$sb.Append("<row r=`"$rowIdx`">")
    for ($c = 0; $c -lt $columns.Count; $c++) {
        $col = $columns[$c]
        $ref = (ColLetter ($c+1)) + $rowIdx
        $key = $col.key
        $val = $row.$key
        $styleAttr = if ($col.wrap) { " s=`"1`"" } else { "" }

        if ($col.numeric -and $val -ne $null -and $val -ne "") {
            [void]$sb.Append("<c r=`"$ref`"$styleAttr><v>$val</v></c>")
        }
        elseif ($col.hyperlink -and $val) {
            $displayKey = "${key}_text"
            $display = $row.$displayKey
            if (-not $display) { $display = $val }
            $si = Get-StringIndex (Esc $display)
            [void]$sb.Append("<c r=`"$ref`" t=`"s`"$styleAttr><v>$si</v></c>")
            $rid = "rIdHL$hlRelId"
            $hlRelId++
            $hlRels.Add("<Relationship Id=`"$rid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`" Target=`"$(Esc $val)`" TargetMode=`"External`"/>") | Out-Null
            $hyperlinkEntries.Add("<hyperlink ref=`"$ref`" r:id=`"$rid`"/>") | Out-Null
        }
        else {
            $text = if ($null -eq $val) { "" } else { [string]$val }
            $si = Get-StringIndex (Esc $text)
            [void]$sb.Append("<c r=`"$ref`" t=`"s`"$styleAttr><v>$si</v></c>")
        }
    }
    [void]$sb.Append("</row>")
    $rowIdx++
}

$lastRowNum = $rowIdx - 1
$lastColLetter = ColLetter $columns.Count
$dimension = "A1:$lastColLetter$lastRowNum"

# column widths
$colsXml = New-Object System.Text.StringBuilder
[void]$colsXml.Append("<cols>")
for ($c = 0; $c -lt $columns.Count; $c++) {
    $w = if ($columns[$c].width) { $columns[$c].width } else { 15 }
    [void]$colsXml.Append("<col min=`"$($c+1)`" max=`"$($c+1)`" width=`"$w`" customWidth=`"1`"/>")
}
[void]$colsXml.Append("</cols>")

$hyperlinksXml = ""
if ($hyperlinkEntries.Count -gt 0) {
    $hyperlinksXml = "<hyperlinks>" + ($hyperlinkEntries -join "") + "</hyperlinks>"
}

$sheetXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="$dimension"/>
<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
$colsXml
<sheetData>
$($sb.ToString())
</sheetData>
<autoFilter ref="$dimension"/>
$hyperlinksXml
</worksheet>
"@

# ---- shared strings xml ----
$ssb = New-Object System.Text.StringBuilder
[void]$ssb.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>")
[void]$ssb.Append("<sst xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" count=`"$($sharedStrings.Count)`" uniqueCount=`"$($sharedStrings.Count)`">")
foreach ($s in $sharedStrings) {
    [void]$ssb.Append("<si><t xml:space=`"preserve`">$s</t></si>")
}
[void]$ssb.Append("</sst>")

# ---- styles.xml (0=default,1=wrap,2=header bold+fill) ----
$stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="10"/><name val="Calibri"/></font>
<font><sz val="10"/><name val="Calibri"/><b/><color rgb="FFFFFFFF"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
</cellXfs>
</styleSheet>
"@

$workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="$(Esc $sheetName)" sheetId="1" r:id="rId1"/></sheets>
</workbook>
"@

$workbookRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

$contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"@

$rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@

$sheetRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
$($hlRels -join "`n")
</Relationships>
"@

# ---- write files to temp dir then zip ----
$tmp = Join-Path $env:TEMP ("xlsxbuild_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path "$tmp\xl\worksheets" | Out-Null
New-Item -ItemType Directory -Force -Path "$tmp\xl\_rels" | Out-Null
New-Item -ItemType Directory -Force -Path "$tmp\_rels" | Out-Null
if ($hlRels.Count -gt 0) { New-Item -ItemType Directory -Force -Path "$tmp\xl\worksheets\_rels" | Out-Null }

[System.IO.File]::WriteAllText("$tmp\[Content_Types].xml", $contentTypes, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText("$tmp\_rels\.rels", $rootRels, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText("$tmp\xl\workbook.xml", $workbookXml, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText("$tmp\xl\_rels\workbook.xml.rels", $workbookRels, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText("$tmp\xl\styles.xml", $stylesXml, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText("$tmp\xl\sharedStrings.xml", $ssb.ToString(), [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText("$tmp\xl\worksheets\sheet1.xml", $sheetXml, [System.Text.UTF8Encoding]::new($false))
if ($hlRels.Count -gt 0) {
    [System.IO.File]::WriteAllText("$tmp\xl\worksheets\_rels\sheet1.xml.rels", $sheetRels, [System.Text.UTF8Encoding]::new($false))
}

if (Test-Path $OutputXlsx) { Remove-Item -Force $OutputXlsx }
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $OutputXlsx, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Remove-Item -Recurse -Force $tmp

Write-Output "OK: $OutputXlsx ($($rows.Count) eiluciu, $($columns.Count) stulpeliu)"
