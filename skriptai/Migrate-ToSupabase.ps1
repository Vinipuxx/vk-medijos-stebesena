<#
.SYNOPSIS
  One-time migration: pushes data/sources.json, data/problems.json (+ occurrences)
  and data/meta.json into the Supabase project defined in .env.
  Safe to re-run: clears sources/problems/occurrences/meta rows first, then reinserts.
  Requires the schema from skriptai/supabase_schema_LOCAL_ONLY.sql to already exist.
#>
$ErrorActionPreference = "Stop"
$root = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena"
$env_ = & "$root\skriptai\Load-Env.ps1"
$baseUrl = $env_["SUPABASE_URL"].TrimEnd("/")
$secret = $env_["SUPABASE_SECRET_KEY"]
if (-not $baseUrl -or -not $secret) { throw "SUPABASE_URL arba SUPABASE_SECRET_KEY trūksta .env faile." }

$headers = @{
    "apikey"        = $secret
    "Authorization" = "Bearer $secret"
    "Content-Type"  = "application/json"
}

function Rest([string]$Method, [string]$Path, $Body, [string]$Prefer = "return=minimal") {
    $h = $headers.Clone()
    $h["Prefer"] = $Prefer
    $uri = "$baseUrl/rest/v1/$Path"
    $json = if ($null -ne $Body) { $Body | ConvertTo-Json -Depth 8 -Compress } else { $null }
    # Supabase's secret key refuses requests whose User-Agent looks like a browser —
    # use a plain script UA so the secret key is accepted (this runs as a local agent, not a browser).
    return Invoke-RestMethod -Method $Method -Uri $uri -Headers $h -Body $json -UserAgent "VK-Mediju-Stebesena-Agent/1.0"
}

Write-Output "== Valau esamas lenteles (occurrences -> problems -> sources) =="
Rest DELETE "occurrences?id=gt.0" $null | Out-Null
Rest DELETE "problems?id=neq.__none__" $null | Out-Null
Rest DELETE "sources?id=neq.00000000-0000-0000-0000-000000000000" $null | Out-Null

Write-Output "== Šaltiniai (data/sources.json) =="
$sources = Get-Content "$root\data\sources.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$sourceRows = $sources | ForEach-Object {
    [PSCustomObject]@{
        nr          = [string]$_.nr
        kategorija  = $_.kategorija
        pavadinimas = $_.pavadinimas
        sritis      = $_.sritis
        url         = $_.url
        monitored   = [bool]$_.monitored
    }
}
Rest POST "sources" $sourceRows | Out-Null
Write-Output "  -> $($sourceRows.Count) šaltinių įkelta"

Write-Output "== Problemos + pasikartojimai (data/problems.json) =="
$problems = Get-Content "$root\data\problems.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$problemRows = @()
$occRows = @()
foreach ($p in $problems) {
    $slug = ($p.title -replace '[^a-zA-Z0-9]+', '-').Trim('-').ToLower()
    if (-not $p.id) { $p | Add-Member -NotePropertyName id -NotePropertyValue $slug -Force }
    $problemRows += [PSCustomObject]@{
        id          = $p.id
        title       = $p.title
        description = $p.description
        sritis      = $p.sritis
        subtema     = $p.subtema
        source_name = $p.sourceName
        source_url  = $p.sourceUrl
        article_date = $p.articleDate
        first_seen  = $p.firstSeen
    }
    $occ = if ($p.occurrences -and $p.occurrences.Count -gt 0) { $p.occurrences } else { @([PSCustomObject]@{ date = $p.articleDate; source = $p.sourceName; url = $p.sourceUrl }) }
    foreach ($o in $occ) {
        $occRows += [PSCustomObject]@{
            problem_id = $p.id
            date       = $o.date
            source     = $o.source
            url        = $o.url
        }
    }
}
Rest POST "problems" $problemRows | Out-Null
Write-Output "  -> $($problemRows.Count) problemų įkelta"
Rest POST "occurrences" $occRows | Out-Null
Write-Output "  -> $($occRows.Count) pasikartojimo įrašų įkelta"

Write-Output "== Metaduomenys (data/meta.json) =="
$meta = Get-Content "$root\data\meta.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$metaBody = [PSCustomObject]@{
    last_run_date   = $meta.lastRunDate
    sources_total   = $meta.sourcesTotal
    sources_checked = $meta.sourcesChecked
    problems_found  = $meta.problemsFound
    scope_note      = $meta.scopeNote
    run_type        = $meta.runType
}
Rest PATCH "meta?id=eq.1" $metaBody | Out-Null
Write-Output "  -> meta atnaujinta"

Write-Output "`nMigracija baigta."
