<#
.SYNOPSIS
  Parses the "Šaltiniai" sheet of the source registry xlsx directly from its OOXML
  (no Excel/python needed) and emits data/sources.json for the dashboard.
#>
param(
    [string]$Xlsx = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena\saltiniai\VK_strateginio_tyrimo_saltiniu_sarasas.xlsx",
    [string]$OutJson = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena\data\sources.json"
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem
$tmp = Join-Path $env:TEMP ("srcparse_" + [guid]::NewGuid().ToString("N"))
[System.IO.Compression.ZipFile]::ExtractToDirectory($Xlsx, $tmp)

function XmlUnescape([string]$s) {
    return $s -replace '&lt;','<' -replace '&gt;','>' -replace '&quot;','"' -replace '&apos;',"'" -replace '&amp;','&'
}
$sstRaw = Get-Content "$tmp\xl\sharedStrings.xml" -Raw -Encoding UTF8
$siBlocks = [regex]::Matches($sstRaw, '<si>(.*?)</si>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
$strings = @($siBlocks | ForEach-Object {
    $inner = $_.Groups[1].Value
    $tMatches = [regex]::Matches($inner, '<t[^>]*>(.*?)</t>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    (($tMatches | ForEach-Object { XmlUnescape $_.Groups[1].Value }) -join "")
})

[xml]$sheet = Get-Content "$tmp\xl\worksheets\sheet2.xml" -Raw -Encoding UTF8
$relsPath = "$tmp\xl\worksheets\_rels\sheet2.xml.rels"
$hyperlinkTargets = @{}
if (Test-Path $relsPath) {
    [xml]$rels = Get-Content $relsPath -Raw -Encoding UTF8
    foreach ($r in $rels.Relationships.Relationship) { $hyperlinkTargets[$r.Id] = $r.Target }
}
$hyperlinkByCell = @{}
if ($sheet.worksheet.hyperlinks) {
    foreach ($h in $sheet.worksheet.hyperlinks.hyperlink) {
        $rid = $h.'r:id'
        if ($rid -and $hyperlinkTargets.ContainsKey($rid)) { $hyperlinkByCell[$h.ref] = $hyperlinkTargets[$rid] }
    }
}

function ColToNum([string]$ref) {
    if ($ref -match '^([A-Z]+)(\d+)$') {
        $letters = $Matches[1]; $n = 0
        foreach ($c in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$c - 64) }
        return $n
    }
    return 0
}

function CellText($cell) {
    if (-not $cell) { return "" }
    if ($cell.t -eq "s") { return $strings[[int]$cell.v] }
    if ($cell.v) { return $cell.v }
    return ""
}

$rows = $sheet.worksheet.sheetData.row
$sources = @()
foreach ($row in $rows) {
    $rowNum = [int]$row.r
    if ($rowNum -le 1) { continue }  # row 1 = column headers
    $cellsByCol = @{}
    foreach ($c in $row.c) {
        $colNum = ColToNum($c.r)
        $cellsByCol[$colNum] = $c
    }
    $colA = CellText $cellsByCol[1]   # Nr.
    $colB = CellText $cellsByCol[2]   # Kategorija (repeated per row within a group)
    $colC = CellText $cellsByCol[3]   # Pavadinimas
    $colD = CellText $cellsByCol[4]   # Sritis
    $colE = CellText $cellsByCol[5]   # URL (display text)

    if (-not $colC) { continue }  # skip blank/instruction rows

    $urlCellRef = $cellsByCol[5].r
    $url = ""
    if ($urlCellRef -and $hyperlinkByCell.ContainsKey($urlCellRef)) { $url = $hyperlinkByCell[$urlCellRef] }
    elseif ($colE -match '^https?://') { $url = $colE }
    if (-not $url) { continue }  # not an actual source row (e.g. a note row)

    $sources += [PSCustomObject]@{
        nr          = $colA
        kategorija  = $colB
        pavadinimas = $colC
        sritis      = $colD
        url         = $url
        monitored   = $true
    }
}

$sources | ConvertTo-Json -Depth 4 | Out-File -FilePath $OutJson -Encoding utf8
Remove-Item -Recurse -Force $tmp
Write-Output "Parsed $($sources.Count) sources -> $OutJson"
