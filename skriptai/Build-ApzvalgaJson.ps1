<#
.SYNOPSIS
  Skaito problemas + pasikartojimus tiesiai iš Supabase (viešas, tik-skaitymo
  publishable raktas) ir sugeneruoja Excel ataskaitą į ataskaitos/.
#>
param(
    [string]$OutXlsx = ""  # jei tuščia, naudojama ataskaitos/VK_zin_apzvalga_<šiandienos data>.xlsx
)

$ErrorActionPreference = "Stop"
$root = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena"
if (-not $OutXlsx) {
    $today = Get-Date -Format "yyyy-MM-dd"
    $OutXlsx = Join-Path $root "ataskaitos\VK_zin_apzvalga_$today.xlsx"
}

$env_ = & "$root\skriptai\Load-Env.ps1"
$baseUrl = $env_["SUPABASE_URL"].TrimEnd("/")
$pubKey = $env_["SUPABASE_PUBLISHABLE_KEY"]
$headers = @{ "apikey" = $pubKey; "Authorization" = "Bearer $pubKey" }

$problems = Invoke-RestMethod -Uri "$baseUrl/rest/v1/problems?select=*,occurrences(date,source,url)&order=sritis,article_date" -Headers $headers

$rows = @()
$i = 1
foreach ($doc in $problems) {
    $occ = @($doc.occurrences | Sort-Object date)
    if ($occ.Count -gt 1) {
        $repeatsTxt = ($occ | ForEach-Object { "$($_.date) | $($_.source) | $($_.url)" }) -join "  ;  "
    } else {
        $repeatsTxt = "—"
    }
    $rows += [PSCustomObject]@{
        nr          = $i
        title       = $doc.title
        desc        = $doc.description
        sourceName  = $doc.source_name
        url         = $doc.source_url
        url_text    = $doc.source_name + " straipsnis"
        date        = $doc.article_date
        sritis      = $doc.sritis
        subtema     = $doc.subtema
        repeats     = $repeatsTxt
        repeatCount = [Math]::Max($occ.Count, 1)
        firstSeen   = $doc.first_seen
    }
    $i++
}

$spec = [PSCustomObject]@{
    sheetName = "Zin. apzvalga"
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

$tmpJson = Join-Path $env:TEMP ("apzvalga_spec_" + [guid]::NewGuid().ToString("N") + ".json")
$spec | ConvertTo-Json -Depth 6 | Out-File -FilePath $tmpJson -Encoding utf8

& "$PSScriptRoot\New-Xlsx.ps1" -InputJson $tmpJson -OutputXlsx $OutXlsx
Remove-Item -Force $tmpJson
Write-Output "Eilučių: $($rows.Count) -> $OutXlsx"
