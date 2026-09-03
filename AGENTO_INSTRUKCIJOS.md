# VK medijų stebėsenos agentas — savaitinės darbo eigos aprašas

Šis dokumentas yra SOP (standartinė darbo tvarka), pagal kurią kas savaitę atliekama
žiniasklaidos ir viešųjų šaltinių stebėsena bei pildoma probleminių temų duomenų bazė.
Jį naudoja tiek žmogus analitikas, tiek Claude (rankiniu būdu arba per suplanuotą
(scheduled) užduotį).

## 1. Kur kas guli

| Kas | Kur |
|---|---|
| Šaltinių registras (96 šaltiniai, 8 grupės) | [saltiniai/VK_strateginio_tyrimo_saltiniu_sarasas.xlsx](saltiniai/VK_strateginio_tyrimo_saltiniu_sarasas.xlsx) |
| Savaitinės žiniasklaidos apžvalgos (Excel, generuojama kas kartą) | `ataskaitos/VK_zin_apzvalga_<data>.xlsx` |
| xlsx generavimo skriptas (be Python/Node — tik .NET/PowerShell) | [skriptai/New-Xlsx.ps1](skriptai/New-Xlsx.ps1) |
| Interaktyvus dašbordas (gyva duomenų bazė + grafikai) | https://claude.ai/code/artifact/3d164a0d-a41a-4b47-b3c4-85a9b4100225 |
| Kanoninis duomenų šaltinis | Dašordo `db` duomenų bazė, kolekcija `problems` + `meta/run` |

**Svarbu:** duomenų bazė dašborde yra vienintelis "šaltinis šaltinių" (source of truth).
Excel `ataskaitos/` aplanke visada regeneruojamas IŠ tos pačios duomenų bazės (arba per
dašbordo "Eksportuoti į .xlsx" mygtuką), kad abu formatai niekada neišsiskirtų.

## 2. Viešojo valdymo sričių sąrašas (16, fiksuotas)

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
(subtema) — trumpas, nuoseklus klasifikatorius srities viduje (pvz. "Pedagogų trūkumas
ir kaita", "Kelių infrastruktūros finansavimas"). Potemes reikia stengtis pakartotinai
naudoti tarp savaičių (ne kurti naują formuluotę tai pačiai problemai) — prieš sukuriant
naują potemę, PATIKRINTI ar panaši jau yra `problems` kolekcijoje.

## 3. Savaitinio paleidimo žingsniai

1. **Šaltinių sąrašas.** Atsidaryti `saltiniai/VK_strateginio_tyrimo_saltiniu_sarasas.xlsx`,
   lapą „Šaltiniai“. Jei stulpelis J („Įtraukta į stebėseną“) užpildytas — tikrinti tik
   eilutes su „Taip“. Jei tuščias — tikrinti visus registre esančius šaltinius (arba
   suderintą pogrupį, jei taip susitarta su analitiku).
2. **Naujienų rinkimas.** Kiekvienam šaltiniui: patikrinti naujienų srautą / naujausius
   straipsnius nuo paskutinio paleidimo datos (žr. `meta/run.lastRunDate` dašborde).
   Naudoti WebSearch (tikslinės užklausos pagal 16 sričių raktažodžius + `site:<domenas>`)
   ir WebFetch (pilnam straipsnio tekstui su citatomis).
3. **Atranka.** Straipsnis įtraukiamas tik jei jis aprašo KONKREČIĄ problemą, riziką,
   trūkumą, pažeidimą ar sisteminę spragą viešajame sektoriuje — ne bendrą naujieną be
   problemos požymio.
4. **Faktų išskyrimas.** Iš straipsnio ištraukti 3–6 TIKSLIUS (žodis į žodį, kabutėse)
   faktus/citatas, kurie pagrindžia problemą: skaičiai, terminai, oficialūs pareigūnų/
   ekspertų pasisakymai. JOKIOS interpretacijos ar apibendrinimo savais žodžiais — jei
   reikia konteksto, jis irgi cituojamas, o ne perfrazuojamas.
5. **Pasikartojimo patikra.** Prieš kuriant naują `problems` įrašą — patikrinti, ar
   panaši problema (ta pati sritis + potemė + iš esmės tas pats klausimas) jau yra
   duomenų bazėje. Jei taip:
   - PRIDĖTI naują elementą į esamo dokumento `occurrences` masyvą (`update`, ne `set`,
     kad nebūtų prarasti seni occurrences — reikia perskaityti dokumentą, papildyti masyvą
     kliento pusėje, tada rašyti visą masyvą iš naujo, nes `update` nesujungia masyvų);
   - atnaujinti `articleDate` į naujausią datą (bet PALIKTI originalų `firstSeen`
     nepakeistą);
   - jei reikia, papildyti `description` nauja informacija (pridėti prie esamo teksto,
     nurodant datą skliaustuose, kaip padaryta pavyzdyje „svietimo-pagalba-veluoja“).
   Jei problema nauja — sukurti naują dokumentą `problems/<trumpas-slug>`.
6. **Rašymas į duomenų bazę.** Naudoti Artifact `write_db` (`batch`, jei kelios eilutės
   vienu metu) su dašbordo URL:
   `https://claude.ai/code/artifact/3d164a0d-a41a-4b47-b3c4-85a9b4100225`.
   Dokumento laukai: `title, description, sritis, subtema, sourceName, sourceUrl,
   articleDate, firstSeen, occurrences[{date, source, url}]`.
7. **`meta/run` atnaujinimas.** Po kiekvieno paleidimo — `set` į `meta/run`:
   `lastRunDate` (šiandienos data), `sourcesTotal` (96), `sourcesChecked` (kiek realiai
   patikrinta), `problemsFound` (kiek naujų + atnaujintų šią savaitę), `scopeNote`
   (laisvas komentaras apie apimtį), `runType` ("weekly-scheduled" arba "manual").
8. **Excel regeneravimas.** Perskaityti visą `problems` kolekciją (`read_db`,
   `db_op: "list"`, `out_dir` į laikiną aplanką), paruošti JSON pagal
   `skriptai/apzvalga_2026-09-03.json` pavyzdį, paleisti:
   ```powershell
   .\skriptai\New-Xlsx.ps1 -InputJson ".\skriptai\apzvalga_<data>.json" -OutputXlsx "..\ataskaitos\VK_zin_apzvalga_<data>.xlsx"
   ```
   (arba paprasčiausiai paspausti „Eksportuoti į .xlsx“ dašborde ir failą išsaugoti į
   `ataskaitos/`).
9. **Patikra.** Atsidaryti dašbordą — patikrinti, kad nauji įrašai matomi grafikuose,
   statistikoje ir sąraše.

## 4. Duomenų bazės schema (`db` capability, dašborto artefakte)

```
problems/<slug>          — vienas dokumentas = viena skirtinga problema
  title            string   trumpas pavadinimas
  description      string   faktai/citatos, žodis į žodį
  sritis           string   TIKSLIAI viena iš 16 sričių (žr. §2)
  subtema          string   klasifikatorius srities viduje
  sourceName       string   naujausio paminėjimo šaltinio pavadinimas
  sourceUrl        string   naujausio paminėjimo nuoroda
  articleDate      string   naujausio paminėjimo data (YYYY-MM-DD)
  firstSeen        string   PIRMO paminėjimo data (YYYY-MM-DD) — niekada nekeisti atgal
  occurrences      array    [{date, source, url}, ...] — visi pasikartojimai chronologškai

meta/run                  — vienas dokumentas, paskutinio paleidimo metaduomenys
  lastRunDate, sourcesTotal, sourcesChecked, problemsFound, scopeNote, runType
```

## 5. Metodinė pastaba (iš originalaus šaltinių sąrašo)

> Žiniasklaidos publikacija strateginiame tyrime yra **hipotezės, o ne įrodymo šaltinis**.
> Prieš teikiant temą į audito planą, signalą reikėtų patvirtinti bent vienu instituciniu
> arba statistiniu šaltiniu (6–8 kategorijos šaltinių registre: Valstybės kontrolė, STT,
> VTEK, Seimo kontrolierių įstaiga, Valstybės duomenų agentūra ir pan.).

## 6. 2026-09-03 bandomojo paleidimo apimtis

Šis pirmasis paleidimas apėmė **10 pagrindinių nacionalinių/institucinių šaltinių**
(neapėmė visų 96 — visų šaltinių savaitinis tikrinimas geriau tinka suplanuotai (scheduled)
užduočiai, kuri gali veikti ilgiau be tiesioginio dialogo). Rasta ir į duomenų bazę
įtraukta **7 problemos** (viena su 2 pasikartojimais), apimančios 4 iš 16 sričių.
Norint įjungti automatinį savaitinį paleidimą per visus 96 šaltinius, žr. pagrindinio
pokalbio santrauką arba paprašyti sukurti suplanuotą (scheduled) užduotį pagal šį
dokumentą.
