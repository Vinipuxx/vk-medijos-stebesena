# VK medijų stebėsenos agentas — savaitinės darbo eigos aprašas

Šis dokumentas yra SOP (standartinė darbo tvarka), pagal kurią kas savaitę atliekama
žiniasklaidos ir viešųjų šaltinių stebėsena bei pildoma probleminių temų duomenų bazė.
Jį naudoja tiek žmogus analitikas, tiek Claude (rankiniu būdu arba per suplanuotą
(scheduled) užduotį).

**Projekto aplankas:** `C:\Users\vmichalovska\Documents\VK_medijos_stebesena`
**Svetainė (GitHub Pages):** https://vinipuxx.github.io/vk-medijos-stebesena/
**Repozitorija:** https://github.com/Vinipuxx/vk-medijos-stebesena
**Duomenų bazė:** Supabase (Postgres), projekto ID `umffyalovfqbajzrrdgu` — raktai `.env` faile.

## 1. Architektūra

```
saltiniai/*.xlsx (istorinis)          Supabase (Postgres) — KANONINIS duomenų šaltinis
                                            problems, occurrences, meta, sources
                                                      │
Agentas (WebSearch/WebFetch + analizė) ──▶ rašo per REST API (.env SUPABASE_SECRET_KEY,
                                            apeina RLS — pilnos teisės)
                                                      │
                        ┌─────────────────────────────┴─────────────────────────────┐
                        ▼                                                             ▼
              index.html (dašbordas, naršyklėje kalbasi              skriptai/Build-ApzvalgaJson.ps1
              su Supabase per @supabase/supabase-js,                          │
              PUBLISHABLE raktu — tik skaito, realaus laiko                    ▼
              atnaujinimai per postgres_changes)                    ataskaitos/*.xlsx (Excel eksportas)

Kodo pakeitimai (index.html, skriptai/ ir pan.) įkeliami: git add -A && git commit && git push
→ GitHub Pages atsinaujina automatiškai. DUOMENYS (problemos/šaltiniai) į GitHub NEKELIAMI —
jie gyvena Supabase ir svetainė juos mato iš karto, be jokio push'o.
```

**Svarbu:** Supabase yra vienintelis kanoninis duomenų šaltinis. `index.html` su juo
kalbasi tiesiogiai naršyklėje (viešu, tik-skaitymo PUBLISHABLE raktu — saugu, RLS
apsaugotas). Agentas rašo per SECRET raktą (`.env`, niekada negit'inamas), kuris
apeina RLS. `data/*.json` failai (jei dar yra repo) yra tik ISTORINIS pradinis
duomenų rinkinys, migruotas per `skriptai/Migrate-ToSupabase.ps1` — jų nebereikia
atnaujinti kas savaitę.

## 2. Kur kas guli

| Kas | Kur |
|---|---|
| Šaltinių registras — istorinis originalas | `saltiniai/VK_strateginio_tyrimo_saltiniu_sarasas.xlsx` |
| **Duomenų bazė (Supabase)** — kanoninis šaltinis viskam | lentelės `sources`, `problems`, `occurrences`, `meta` |
| Prisijungimo duomenys prie Supabase (NEGIT'inama) | `.env` (šablonas: `.env.example`) |
| `.env` skaitymo pagalbinis skriptas | `skriptai/Load-Env.ps1` |
| Vienkartinė migracija JSON → Supabase (jau panaudota) | `skriptai/Migrate-ToSupabase.ps1` |
| Duomenų bazės schemos SQL (jau paleista, NEGIT'inama — turi PIN atviru tekstu) | `skriptai/supabase_schema_LOCAL_ONLY.sql` |
| Dašbordas (naršyklėje kalbasi su Supabase per JS klientą) | `index.html` |
| Excel eksportas (generuojamas TIESIAI iš Supabase) | `ataskaitos/VK_zin_apzvalga_<data>.xlsx` |
| xlsx generavimo skriptas (be Python/Node — tik .NET/PowerShell) | `skriptai/New-Xlsx.ps1`, `skriptai/Build-ApzvalgaJson.ps1` |
| Šio paleidimo detalus žurnalas (kokie šaltiniai tikrinti, kas rasta/nerasta) | `skriptai/pilnas_patikrinimo_zurnalas_<data>.md` (kurti naują kiekvienam pilnam paleidimui) |

## 3. Viešojo valdymo sričių sąrašas (16, fiksuotas)

1. Aplinka, miškai ir klimato kaita
2. Ekonomikos konkurencingumas
3. Energetika
4. Informacinių išteklių valdymas
5. Kultūra ir visuomenės informavimas
6. Socialinė apsauga ir užimtumas
7. Sveikatos apsauga
8. Švietimas, mokslas ir sportas
9. Teisingumas
10. Transportas ir ryšiai
11. Užsienio politika
12. Valstybės saugumas ir gynyba
13. Viešasis saugumas
14. Viešasis valdymas, regioninė politika ir viešasis administravimas
15. Viešieji finansai ir oficialioji statistika
16. Žemės ir maisto ūkis, kaimo plėtra, žuvininkystė, veterinarija ir žemės tvarkymas

Kiekvienai nustatytai problemai priskiriama **lygiai viena** sritis iš šio sąrašo (pagal
straipsnio turinį, ne pagal šaltinio kategoriją registre) + laisvos formos **potemė**
(subtema) — trumpas, nuoseklus klasifikatorius srities viduje (pvz. „Pedagogų trūkumas
ir kaita", „Kelių infrastruktūros finansavimas"). Potemes reikia stengtis pakartotinai
naudoti tarp savaičių (ne kurti naują formuluotę tai pačiai problemai) — prieš sukuriant
naują potemę, PATIKRINTI `data/problems.json`, ar panaši jau yra.

## 4. Savaitinio paleidimo žingsniai

Visi duomenų veiksmai vyksta per Supabase REST API (PostgREST), naudojant SECRET raktą
(apeina RLS — pilnos teisės skaityti/rašyti). Pavyzdys, kaip gauti antraštes PowerShell'e:

```powershell
$env_ = & "C:\Users\vmichalovska\Documents\VK_medijos_stebesena\skriptai\Load-Env.ps1"
$base = $env_["SUPABASE_URL"].TrimEnd("/")
$h = @{ "apikey" = $env_["SUPABASE_SECRET_KEY"]; "Authorization" = "Bearer $($env_["SUPABASE_SECRET_KEY"])"; "Content-Type" = "application/json" }
```

1. **Pradinė būsena.**
   ```powershell
   $problems = Invoke-RestMethod "$base/rest/v1/problems?select=*,occurrences(date,source,url)" -Headers $h
   $meta = Invoke-RestMethod "$base/rest/v1/meta?id=eq.1&select=*" -Headers $h
   ```
   `$meta.last_run_date` sako, nuo kada ieškoti naujienų. `$problems` — jau žinomos temos
   (reikės tikrinti pasikartojimus).
2. **Šaltinių sąrašas.**
   ```powershell
   $sources = Invoke-RestMethod "$base/rest/v1/sources?monitored=eq.true&select=*" -Headers $h
   ```
   Tai KANONINIS, gyvas šaltinių sąrašas — svetainės lankytojai gali jį redaguoti per
   „Stebimi šaltiniai" skydelį (su PIN kodu) realiu laiku, tad visada skaityk jį iš
   naujo per kiekvieną paleidimą, nesiremk sena kopija.
3. **Naujienų rinkimas.** Kiekvienam šaltiniui: WebFetch homepage/naujienų srautą,
   nustatyti ar yra naujų straipsnių nuo `meta.last_run_date`. Naudoti WebSearch
   tiksliniams raktažodžiams pagal 16 sričių, kai homepage tiesioginis nuskaitymas
   blokuojamas (žinoma, kad Delfi, Lrytas, ELTA, Alfa, Verslo žinios WebFetch blokuoja —
   403/klaida; šiems reikėtų arba RSS adreso, arba praleisti su pastaba).
4. **Atranka.** Straipsnis įtraukiamas tik jei aprašo KONKREČIĄ problemą, riziką,
   trūkumą, pažeidimą ar sisteminę spragą viešajame sektoriuje — ne bendrą naujieną.
5. **Faktų išskyrimas.** Ištraukti 3–6 TIKSLIUS (žodis į žodį, kabutėse) faktus/citatas:
   skaičiai, terminai, oficialūs pasisakymai. JOKIOS interpretacijos ar apibendrinimo
   savais žodžiais.
6. **Pasikartojimo patikra ir rašymas.** Palyginti su `$problems`: jei panaši problema
   (ta pati sritis + potemė + iš esmės tas pats klausimas) jau yra —
   ```powershell
   # naujas pasikartojimas jau žinomai problemai:
   Invoke-RestMethod "$base/rest/v1/occurrences" -Method Post -Headers $h -Body (
     @{ problem_id = "<esamas-id>"; date = "2026-09-08"; source = "..."; url = "..." } | ConvertTo-Json)
   # jei reikia atnaujinti aprašymą/naujausią datą:
   Invoke-RestMethod "$base/rest/v1/problems?id=eq.<esamas-id>" -Method Patch -Headers $h -Body (
     @{ description = "<papildytas tekstas>"; article_date = "2026-09-08" } | ConvertTo-Json)
   ```
   PALIKTI `first_seen` NEPAKEISTĄ. Jei problema nauja:
   ```powershell
   $newId = ($title -replace '[^a-zA-Z0-9]+','-').Trim('-').ToLower()
   Invoke-RestMethod "$base/rest/v1/problems" -Method Post -Headers $h -Body (
     @{ id = $newId; title = $title; description = $desc; sritis = $sritis; subtema = $subtema
        source_name = $srcName; source_url = $srcUrl; article_date = $date; first_seen = $date
      } | ConvertTo-Json)
   Invoke-RestMethod "$base/rest/v1/occurrences" -Method Post -Headers $h -Body (
     @{ problem_id = $newId; date = $date; source = $srcName; url = $srcUrl } | ConvertTo-Json)
   ```
7. **`meta` atnaujinimas.**
   ```powershell
   Invoke-RestMethod "$base/rest/v1/meta?id=eq.1" -Method Patch -Headers $h -Body (
     @{ last_run_date = (Get-Date -Format "yyyy-MM-dd"); sources_total = 96
        sources_checked = $N; problems_found = $problems.Count; run_type = "weekly-scheduled"
        scope_note = "..." } | ConvertTo-Json)
   ```
8. **Excel regeneravimas** (skaito tiesiai iš Supabase, publishable raktu):
   ```powershell
   & "C:\Users\vmichalovska\Documents\VK_medijos_stebesena\skriptai\Build-ApzvalgaJson.ps1"
   ```
9. **Žurnalas.** Parašyti/atnaujinti `skriptai/pilnas_patikrinimo_zurnalas_<data>.md` —
   kokie šaltiniai tikrinti, kas rasta, kas blokuota/nepasiekiama.
10. **Kodo/ataskaitos įkėlimas į GitHub** (duomenys jau Supabase, čia keliasi tik naujas
    Excel failas ir žurnalas):
    ```bash
    cd "C:\Users\vmichalovska\Documents\VK_medijos_stebesena"
    git add ataskaitos skriptai/pilnas_patikrinimo_zurnalas_*.md
    git commit -m "Savaitinis atnaujinimas <data>: N naujų/atnaujintų problemų"
    git push
    ```
    (`git add -A` NENAUDOTI — tai galėtų netyčia įkelti `.env` ar `*_LOCAL_ONLY.sql`,
    nors jie ir `.gitignore` sąraše, saugiau likti eksplicitiškam.)
11. **Patikra.** Atidaryti https://vinipuxx.github.io/vk-medijos-stebesena/ — duomenys
    ten atsinaujina IŠ KARTO po Supabase rašymo (realaus laiko prenumerata), push
    reikalingas tik svetainės kodui/Excel failui, ne patiems duomenims.

## 5. Duomenų bazės schema (Supabase)

**`problems`** — viena eilutė = viena skirtinga problema (`id` = trumpas slug):
`id, title, description, sritis, subtema, source_name, source_url, article_date, first_seen`

**`occurrences`** — visi paminėjimai (vienas ar daugiau per problemą):
`id (auto), problem_id (FK -> problems.id), date, source, url`

**`meta`** — viena eilutė (`id = 1`):
`last_run_date, sources_total, sources_checked, problems_found, scope_note, run_type`

**`sources`** — šaltinių registras, redaguojamas per svetainę (PIN apsauga):
`id (uuid), nr, kategorija, pavadinimas, sritis, url, monitored`

Pilna schema su RLS politikomis ir PIN funkcijomis: `skriptai/supabase_schema_LOCAL_ONLY.sql`
(NEGIT'inamas — turi PIN atviru tekstu; jei reikia peržiūrėti, atsidaryti lokaliai).

## 6. Metodinė pastaba (iš originalaus šaltinių sąrašo)

> Žiniasklaidos publikacija strateginiame tyrime yra **hipotezės, o ne įrodymo šaltinis**.
> Prieš teikiant temą į audito planą, signalą reikėtų patvirtinti bent vienu instituciniu
> arba statistiniu šaltiniu (6–8 kategorijos šaltinių registre: Valstybės kontrolė, STT,
> VTEK, Seimo kontrolierių įstaiga, Valstybės duomenų agentūra ir pan.).

## 7. Istorija

- **2026-09-03** — pirmas bandomasis paleidimas (10 šaltinių, 7 problemos), po to pilnas
  paleidimas (68/96 šaltinių, 13 problemų). Nuo Claude Artifact (gyva duomenų bazė,
  reikėjo prisijungimo) pereita prie statinio GitHub Pages sprendimo su
  `data/problems.json`. Sukurtas savaitinis automatinis paleidimas (pirmadieniais 8:00).
  Pridėtas `data/sources.json` — šaltinių sąrašas dabar redaguojamas per svetainę
  („Stebimi šaltiniai" skydelis), o ne tik per originalų Excel registrą.
- **2026-09-03 (vėliau tą pačią dieną)** — pereita nuo statinių JSON failų prie
  **Supabase** (Postgres) kaip vienintelio kanoninio duomenų šaltinio. Svetainė dabar
  realiu laiku kalbasi su duomenų baze (`@supabase/supabase-js`, publishable raktas).
  Šaltinių redagavimas per svetainę dabar TIKRAI persistuoja visiems (ne tik
  naršyklės localStorage) — apsaugotas PIN kodu per Postgres `SECURITY DEFINER`
  funkcijas. Agentas rašo per atskirą secret raktą (`.env`, apeina RLS).
