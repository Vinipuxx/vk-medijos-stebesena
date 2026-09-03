<#
  Savaitinio (papildomo) paleidimo įrašymas į Supabase - 2026-09-03, antra banga
  (rajoniniai laikraščiai, kurie nebuvo patikrinti pirmame pilname paleidime).
#>
$ErrorActionPreference = "Stop"
$root = "C:\Users\vmichalovska\Documents\VK_medijos_stebesena"
$env_ = & "$root\skriptai\Load-Env.ps1"
$baseUrl = $env_["SUPABASE_URL"].TrimEnd("/")
$secret = $env_["SUPABASE_SECRET_KEY"]
$headers = @{ "apikey" = $secret; "Authorization" = "Bearer $secret"; "Content-Type" = "application/json" }

function Rest([string]$Method, [string]$Path, $Body, [string]$Prefer = "return=representation") {
    $h = $headers.Clone(); $h["Prefer"] = $Prefer
    $json = if ($null -ne $Body) { $Body | ConvertTo-Json -Depth 8 -Compress } else { $null }
    return Invoke-RestMethod -Method $Method -Uri "$baseUrl/rest/v1/$Path" -Headers $h -Body $json -UserAgent "VK-Mediju-Stebesena-Agent/1.0"
}

# --- 1) Nauja problema: elektros tinklo atkūrimas po audros ---
$p1 = [PSCustomObject]@{
    id = "elektros-tinklo-atkurimas-po-audros"
    title = "Po galingiausios nuo 1970 m. audros elektros tinklas liko sugriautas, savivaldybių vadovai vengia komentuoti atkūrimo eigą"
    description = '"Šio ryto duomenimis, šiek tiek virš 4 tūkst. [vartotojų] skirtingose šalies vietose neturi elektros energijos." "Blogiausia situacija ir toliau išlieka Utenos regione, taip pat Anykščių, Zarasų apylinkėse." "Tinklas iš tikrųjų yra sugriautas. Šita audra... buvo galingiausia nuo 1970 metų." "Kai kuriose vietovėse elektros tinklas yra taip stipriai pažeistas, jog šį reikės tiesti iš naujo." "Piko metu elektros tiekimas buvo sutrikęs daugiau kaip 268 tūkst. ESO klientų." Anykščių rajone užregistruoti "243 elektros tiekimo sutrikimai su 759 atjungtų klientų", atkūrimui dirbo "28 brigados... 11 skirtingų Anykščių rajono seniūnijų". Anykščių meras Kęstutis Tubis viešai atsakė trumpai: "Nebeturiu komentarų!!!"'
    sritis = "Energetika"
    subtema = "Elektros tinklų atkūrimas po stichinės nelaimės ir savivaldybių komunikacija"
    source_name = "Anykšta"
    source_url = "https://www.anyksta.lt/eso-po-audros-elektros-vis-dar-neturi-per-4-tukst-vartotoju/"
    article_date = "2026-09-03"
    first_seen = "2026-09-03"
}
Rest POST "problems" $p1 "resolution=merge-duplicates,return=minimal" | Out-Null
$occ1 = @(
  [PSCustomObject]@{ problem_id = $p1.id; date = "2026-09-03"; source = "Anykšta (ESO pranešimas)"; url = "https://www.anyksta.lt/eso-po-audros-elektros-vis-dar-neturi-per-4-tukst-vartotoju/" },
  [PSCustomObject]@{ problem_id = $p1.id; date = "2026-09-03"; source = "Anykšta (mero komentaras)"; url = "https://www.anyksta.lt/meras-apie-elektra-nebeturiu-komentaru/" }
)
Rest POST "occurrences" $occ1 | Out-Null
Write-Output "OK: $($p1.id)"

# --- 2) Nauja problema: Stasys Museum lankytojų mažėjimas / biudžeto lėšos ---
$p2 = [PSCustomObject]@{
    id = "stasys-museum-lankytoju-mazejimas"
    title = "Panevėžio savivaldybės finansuojamas 'Stasys Museum' praranda lankytojus, savų pajamų dalis lieka minimali"
    description = '"Pernai „Stasys Museum" apsilankė 70 591 lankytojas – tai net 12 239 lankytojais mažiau nei užpernai." Finansavimas: "beveik 2,3 mln. eurų – sudarė Panevėžio miesto savivaldybės biudžeto pinigai." "Muziejus pernai iš parduotų bilietų ir suteiktų paslaugų užsidirbo vos 114,5 tūkst. eurų, arba tik 4,35 procento." "Šių metų pirmąjį pusmetį „Stasys Museum" aplankė 34 645 lankytojai, vidutiniškai per mėnesį – beveik po 5800." Palyginimui, "per pirmąjį veiklos mėnesį užpraėjusią vasarą muziejus sulaukė beveik 20 tūkst. lankytojų."'
    sritis = "Kultūra ir visuomenės informavimas"
    subtema = "Viešai finansuojamų kultūros įstaigų lankomumas ir savų pajamų dalis"
    source_name = "Panevėžio kraštas"
    source_url = "https://panskliautas.lt/kulturos-stebuklas-vis-labiau-bliuksta"
    article_date = "2026-08-28"
    first_seen = "2026-08-28"
}
Rest POST "problems" $p2 "resolution=merge-duplicates,return=minimal" | Out-Null
$occ2 = @([PSCustomObject]@{ problem_id = $p2.id; date = "2026-08-28"; source = "Panevėžio kraštas"; url = "https://panskliautas.lt/kulturos-stebuklas-vis-labiau-bliuksta" })
Rest POST "occurrences" $occ2 | Out-Null
Write-Output "OK: $($p2.id)"

# --- 3) meta atnaujinimas ---
$metaBody = [PSCustomObject]@{
    last_run_date = "2026-09-03"
    sources_total = 96
    sources_checked = 82
    problems_found = 15
    scope_note = "Antra banga: patikrinti dar 14 rajoninių laikraščių (2 nauji signalai rasti), iš viso 82/96 šaltinių patikrinta"
    run_type = "manual"
}
Rest PATCH "meta?id=eq.1" $metaBody "return=minimal" | Out-Null
Write-Output "OK: meta"
