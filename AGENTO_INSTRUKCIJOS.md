# VK medijų stebėsenos agentas — savaitinės darbo eigos aprašas

Šis dokumentas yra SOP (standartinė darbo tvarka), pagal kurią kas savaitę atliekama
žiniasklaidos ir viešųjų šaltinių stebėsena bei pildoma probleminių temų duomenų bazė.
Jį naudoja tiek žmogus analitikas, tiek Claude (rankiniu būdu arba per suplanuotą
(scheduled) užduotį).

**Projekto aplankas:** `C:\Users\vmichalovska\Documents\VK_medijos_stebesena`
**Svetainė (GitHub Pages):** https://vinipuxx.github.io/vk-medijos-stebesena/
**Repozitorija:** https://github.com/Vinipuxx/vk-medijos-stebesena

## 1. Architektūra

```
saltiniai/*.xlsx  ──▶  Agentas (WebSearch/WebFetch + analizė)
                              │
                              ▼
                   data/problems.json   ◀── KANONINIS duomenų šaltinis („lentelė")
                   data/meta.json       ◀── paskutinio paleidimo metaduomenys
                              │
                ┌─────────────┴─────────────┐
                ▼                             ▼
        index.html (dašbordas,        skriptai/Build-ApzvalgaJson.ps1
        fetch'ina JSON, jokio                  │
        prisijungimo nereikia)                 ▼
                                     ataskaitos/*.xlsx (Excel eksportas)

Viskas įkeliama: git add -A && git commit && git push  →  GitHub Pages atsinaujina automatiškai
```

**Svarbu:** `data/problems.json` yra vienintelis kanoninis šaltinis. `index.html` jį
tiesiog nuskaito naršyklėje (`fetch`) — jokios duomenų bazės, prisijungimo ar serverio
nereikia. `ataskaitos/*.xlsx` visada regeneruojamas IŠ `data/problems.json`
(`skriptai/Build-ApzvalgaJson.ps1`), kad Excel ir svetainė niekada neišsiskirtų.

## 2. Kur kas guli

| Kas | Kur |
|---|---|
| Šaltinių registras — istorinis originalas | `saltiniai/VK_strateginio_tyrimo_saltiniu_sarasas.xlsx` |
| **Kanoninis, gyvas šaltinių sąrašas** (redaguojamas per svetainę) | `data/sources.json` |
| **Kanoniniai duomenys** (visos nustatytos problemos) | `data/problems.json` |
| Paskutinio paleidimo metaduomenys | `data/meta.json` |
| Dašbordas (statinis, be backend'o) | `index.html` |
| Excel eksportas (generuojamas iš data/problems.json) | `ataskaitos/VK_zin_apzvalga_<data>.xlsx` |
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

1. **Pradinė būsena.** Perskaityti `data/problems.json` (esami įrašai — reikės tikrinti
   pasikartojimus) ir `data/meta.json` (`lastRunDate` — nuo kada ieškoti naujų straipsnių).
2. **Šaltinių sąrašas.** Skaityti `data/sources.json` — tai KANONINIS, gyvas šaltinių
   sąrašas (jį svetainės lankytojai gali redaguoti per „Stebimi šaltiniai" skydelį ir
   atsiųsti atnaujintą versiją; jei tokia gauta, ji jau bus pakeitusi šį failą repo).
   Tikrinti tik įrašus, kur `"monitored": true`. (`saltiniai/VK_strateginio_tyrimo_saltiniu_sarasas.xlsx`
   lieka istorinis originalas — `data/sources.json` sugeneruotas iš jo per
   `skriptai/Parse-Sources.ps1`, bet nuo šiol redaguojamas per svetainę, ne per Excel.)
3. **Naujienų rinkimas.** Kiekvienam šaltiniui: WebFetch homepage/naujienų srautą,
   nustatyti ar yra naujų straipsnių nuo `lastRunDate`. Naudoti WebSearch tiksliniams
   raktažodžiams pagal 16 sričių, kai homepage tiesioginis nuskaitymas blokuojamas
   (žinoma, kad Delfi, Lrytas, ELTA, Alfa, Verslo žinios WebFetch blokuoja — 403/klaida;
   šiems reikėtų arba RSS adreso iš registro F stulpelio, arba praleisti su pastaba).
4. **Atranka.** Straipsnis įtraukiamas tik jei aprašo KONKREČIĄ problemą, riziką,
   trūkumą, pažeidimą ar sisteminę spragą viešajame sektoriuje — ne bendrą naujieną.
5. **Faktų išskyrimas.** Ištraukti 3–6 TIKSLIUS (žodis į žodį, kabutėse) faktus/citatas:
   skaičiai, terminai, oficialūs pasisakymai. JOKIOS interpretacijos ar apibendrinimo
   savais žodžiais.
6. **Pasikartojimo patikra.** Palyginti su `data/problems.json`: jei panaši problema
   (ta pati sritis + potemė + iš esmės tas pats klausimas) jau yra —
   - PRIDĖTI naują elementą į to įrašo `occurrences` masyvą (chronologiškai);
   - atnaujinti `articleDate` į naujausią datą (PALIKTI `firstSeen` nepakeistą);
   - jei reikia, papildyti `description` nauja informacija (pridėti prie esamo teksto,
     nurodant datą skliaustuose — žr. pavyzdį `svietimo-pagalba-veluoja` faile).
   Jei problema nauja — sukurti naują objektą masyve (žr. schemą §5).
7. **`data/problems.json` perrašymas.** Įrašyti visą atnaujintą masyvą (senus + naujus/
   pakeistus įrašus) atgal į `data/problems.json` (UTF-8, be BOM, gražiai suformatuotą
   JSON — `ConvertTo-Json -Depth 8` PowerShell'e arba lygiavertis).
8. **`data/meta.json` atnaujinimas.** `lastRunDate` (šiandienos data, YYYY-MM-DD),
   `sourcesTotal` (96), `sourcesChecked` (kiek realiai patikrinta šį paleidimą),
   `problemsFound` (bendras `data/problems.json` įrašų skaičius), `scopeNote` (laisvas
   komentaras apie apimtį), `runType` (`"weekly-scheduled"` arba `"manual"`).
9. **Excel regeneravimas.**
   ```powershell
   & "C:\Users\vmichalovska\Documents\VK_medijos_stebesena\skriptai\Build-ApzvalgaJson.ps1"
   ```
10. **Žurnalas.** Parašyti/atnaujinti `skriptai/pilnas_patikrinimo_zurnalas_<data>.md` —
    kokie šaltiniai tikrinti, kas rasta, kas blokuota/nepasiekiama.
11. **Įkėlimas į GitHub.**
    ```bash
    cd "C:\Users\vmichalovska\Documents\VK_medijos_stebesena"
    git add -A
    git commit -m "Savaitinis atnaujinimas <data>: N naujų/atnaujintų problemų"
    git push
    ```
    GitHub Pages svetainė atsinaujina automatiškai per ~1 min.
12. **Patikra.** Atidaryti https://vinipuxx.github.io/vk-medijos-stebesena/ ir
    patikrinti, kad nauji įrašai matomi statistikoje, grafikuose ir sąraše.

## 5. `data/problems.json` schema

Masyvas objektų:

```json
{
  "title": "Trumpas problemos pavadinimas",
  "description": "Faktai/citatos žodis į žodį iš straipsnio (be interpretacijos)",
  "sritis": "TIKSLIAI viena iš 16 sričių (žr. §3)",
  "subtema": "Laisvos formos klasifikatorius srities viduje",
  "sourceName": "Naujausio paminėjimo šaltinio pavadinimas",
  "sourceUrl": "Naujausio paminėjimo nuoroda",
  "articleDate": "YYYY-MM-DD (naujausio paminėjimo data)",
  "firstSeen": "YYYY-MM-DD (PIRMO paminėjimo data — niekada nekeisti atgal)",
  "occurrences": [
    { "date": "YYYY-MM-DD", "source": "Šaltinio pavadinimas", "url": "https://..." }
  ]
}
```

`data/meta.json`:
```json
{
  "lastRunDate": "YYYY-MM-DD",
  "sourcesTotal": 96,
  "sourcesChecked": 68,
  "problemsFound": 13,
  "scopeNote": "Laisvas komentaras",
  "runType": "weekly-scheduled"
}
```

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
