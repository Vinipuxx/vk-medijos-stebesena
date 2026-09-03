<#
.SYNOPSIS
  Konvertuoja skriptai/db_export/problems/*.json (read_db eksportas) į
  New-Xlsx.ps1 laukiamą JSON formatą ir sugeneruoja galutinį .xlsx.
#>
param(
    [string]$ExportDir = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena\skriptai\db_export\problems",
    [string]$OutJson = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena\skriptai\apzvalga_pilna_2026-09-03.json",
    [string]$OutXlsx = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena\ataskaitos\VK_zin_apzvalga_2026-09-03.xlsx"
)

$files = Get-ChildItem -Path $ExportDir -Filter "*.json"
$rows = @()
$i = 1
foreach ($f in ($files | Sort-Object { (Get-Content $_.FullName -Raw | ConvertFrom-Json).sritis }, Name)) {
    $doc = Get-Content $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $occ = $doc.occurrences
    if ($occ.Count -gt 1) {
        $repeatsTxt = ($occ | ForEach-Object { "$($_.date) | $($_.source) | $($_.url)" }) -join "  ;  "
    } else {
        $repeatsTxt = "—"
    }
    $rows += [PSCustomObject]@{
        nr          = $i
        title       = $doc.title
        desc        = $doc.description
        sourceName  = $doc.sourceName
        url         = $doc.sourceUrl
        url_text    = $doc.sourceName + " straipsnis"
        date        = $doc.articleDate
        sritis      = $doc.sritis
        subtema     = $doc.subtema
        repeats     = $repeatsTxt
        repeatCount = $occ.Count
        firstSeen   = $doc.firstSeen
    }
    $i++
}

$spec = [PSCustomObject]@{
    sheetName = "Zin. apzvalga 2026-09-03"
    columns = @(
        @{ header = "Nr."; key = "nr"; width = 5; numeric = $true }
        @{ header = "Problemos pavadinimas"; key = "title"; width = 38; wrap = $true }
        @{ header = "Detalus problemos aprašymas (faktai/citatos iš straipsnio)"; key = "desc"; width = 70; wrap = $true }
        @{ header = "Šaltinis (svetainė)"; key = "sourceName"; width = 16; wrap = $true }
        @{ header = "Nuoroda į straipsnį"; key = "url"; width = 20; hyperlink = $true }
        @{ header = "Straipsnio data"; key = "date"; width = 13; wrap = $true }
        @{ header = "Viešojo valdymo sritis"; key = "sritis"; width = 30; wrap = $true }
        @{ header = "Potemė (klasifikatorius srityje)"; key = "subtema"; width = 34; wrap = $true }
        @{ header = "Pasikartojimai (data | šaltinis | nuoroda)"; key = "repeats"; width = 60; wrap = $true }
        @{ header = "Pasikartojimų skaičius"; key = "repeatCount"; width = 10; numeric = $true }
        @{ header = "Pirmą kartą pastebėta"; key = "firstSeen"; width = 13; wrap = $true }
    )
    rows = $rows
}

$spec | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutJson -Encoding utf8

& "$PSScriptRoot\New-Xlsx.ps1" -InputJson $OutJson -OutputXlsx $OutXlsx
Write-Output "Eilučių: $($rows.Count)"
