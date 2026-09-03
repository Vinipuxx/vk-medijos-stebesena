# VK medijų stebėsenos dašbordas

Savaitinės žiniasklaidos ir viešųjų šaltinių stebėsenos sistema: agentas renka
probleminius signalus iš žiniasklaidos ir institucijų, klasifikuoja juos pagal 16
viešojo valdymo sričių, seka pasikartojimus tarp savaičių/šaltinių/mėnesių/metų, ir
viską pateikia interaktyviame dašborde bei eksportuojamoje Excel lentelėje.

**Svetainė (gyva):** **https://vinipuxx.github.io/vk-medijos-stebesena/**
(GitHub Pages, `Settings → Pages → Deploy from a branch → main / (root)` — jau įjungta.)

**Automatinis atnaujinimas:** suplanuota užduotis `vk-medijos-stebesena-weekly`
paleidžia agentą kiekvieną **pirmadienį 8:00** (veikia, kol Claude programa atidaryta
kompiuteryje; jei uždaryta — pasileidžia kitą kartą atidarius). Rankiniu būdu galima
paleisti bet kada, paprašius Claude sekti `AGENTO_INSTRUKCIJOS.md`.

### Ką rodo dašbordas

- **Statistikos kubeliai** su paaiškinimais, kaip skaičiai susiję vienas su kitu
  (kiek problemų, kiek iš jų pasikartoja, kiek signalų iš viso, kiek sričių paliesta).
- **„Daugiausiai kartojasi"** perspėjimo juosta — iš karto matoma didžiausio prioriteto tema.
- **„Temų svoris"** — visos problemos surikiuotos pagal pasikartojimų skaičių (ne pagal
  sritį, o pagal konkrečią temą), kad būtų matyti, kuri konkreti tema kartojasi dažniausiai.
- **Sričių stulpelinė diagrama** (16 viešojo valdymo sričių) su filtravimu paspaudus.
- **Signalų dinamikos grafikas** su Savaitė / Mėnuo / Metai perjungikliais ir
  „Kaupiamai" (bendra augančia suma) rodymu; užvedus pelę ant taško — suskirstymas pagal sritį.
- **Pilnas temų sąrašas** su paieška, filtrais ir originaliomis citatomis iš straipsnių.
- **Eksportas į .xlsx** vienu mygtuko paspaudimu (standartinis naršyklės atsisiuntimas).

## Architektūra

```
Šaltinių registras (saltiniai/) ──▶ Agentas (Claude, rankinis arba suplanuotas paleidimas)
                                          │  seka AGENTO_INSTRUKCIJOS.md
                                          ▼
                          data/problems.json  +  data/meta.json   ◀── "lentelė"
                                          │
                       ┌──────────────────┴──────────────────┐
                       ▼                                       ▼
              index.html (dašbordas,                  ataskaitos/*.xlsx
              skaito JSON, jokio                       (Excel eksportas,
              prisijungimo nereikia)                    tas pats turinys)
```

`index.html` yra **statinis** puslapis — jis tiesiog nuskaito `data/problems.json`
naršyklėje (`fetch`) ir viską sudėlioja į grafikus/lenteles JavaScript'u. Jokio
serverio, duomenų bazės ar prisijungimo nereikia — bet tai reiškia, kad puslapis
rodo tai, kas buvo `data/problems.json` **paskutinio commit'o** metu. Kad
informacija atsinaujintų, reikia paleisti agentą iš naujo ir įkelti (`git push`)
naują `data/problems.json` versiją.

## Aplankų struktūra

| Aplankas / failas | Paskirtis |
|---|---|
| `index.html` | Dašbordas (grafikai, filtrai, eksportas į .xlsx) |
| `data/problems.json` | Visų nustatytų problemų duomenų bazė (JSON) — tai ir yra „lentelė", iš kurios pildosi dašbordas |
| `data/meta.json` | Paskutinio paleidimo metaduomenys (data, patikrintų šaltinių sk. ir pan.) |
| `saltiniai/` | Šaltinių registras (96 šaltiniai, 8 grupės) |
| `ataskaitos/` | Excel eksportai (ta pati informacija kaip `data/problems.json`, patogu peržiūrai/archyvavimui) |
| `skriptai/` | Pagalbiniai PowerShell skriptai (.xlsx generavimas be Python/Node, JSON konvertavimas) |
| `AGENTO_INSTRUKCIJOS.md` | Pilna savaitinio rinkimo darbo eiga (SOP) — ką agentas turi daryti kas savaitę |
| `dashboard/dashboard.html` | Ankstesnė versija, veikusi kaip Claude Artifact su gyva duomenų baze (archyvas) |

## Kaip atnaujinami duomenys

**Automatiškai:** suplanuota užduotis `vk-medijos-stebesena-weekly` kas pirmadienį
8:00 seka `AGENTO_INSTRUKCIJOS.md` ir pati viską atlieka — nuo šaltinių tikrinimo iki
`git push`. Valdyti (pristabdyti, keisti laiką) galima Claude programos skiltyje „Scheduled".

**Rankiniu būdu:**
1. Paprašykite Claude paleisti stebėseną pagal `AGENTO_INSTRUKCIJOS.md`.
2. Agentas surenka naujus signalus, atnaujina `data/problems.json` ir `data/meta.json`,
   regeneruoja `ataskaitos/*.xlsx` (`skriptai/Build-ApzvalgaJson.ps1`).
3. Pakeitimai įkeliami į šią repozitoriją (`git add`, `git commit`, `git push`).
4. GitHub Pages automatiškai perdiegia svetainę per ~1 min.

## Duomenų schema (`data/problems.json`)

Kiekvienas masyvo elementas:

```json
{
  "title": "Trumpas problemos pavadinimas",
  "description": "Faktai/citatos žodis į žodį iš straipsnio (be interpretacijos)",
  "sritis": "Viena iš 16 viešojo valdymo sričių",
  "subtema": "Laisvos formos klasifikatorius srities viduje",
  "sourceName": "Naujausio paminėjimo šaltinio pavadinimas",
  "sourceUrl": "Naujausio paminėjimo nuoroda",
  "articleDate": "YYYY-MM-DD (naujausio paminėjimo data)",
  "firstSeen": "YYYY-MM-DD (pirmo paminėjimo data)",
  "occurrences": [
    { "date": "YYYY-MM-DD", "source": "Šaltinio pavadinimas", "url": "https://..." }
  ]
}
```

## Metodinė pastaba

> Žiniasklaidos publikacija strateginiame tyrime yra **hipotezės, o ne įrodymo šaltinis**.
> Prieš teikiant temą į audito planą, signalą reikėtų patvirtinti bent vienu instituciniu
> arba statistiniu šaltiniu.
