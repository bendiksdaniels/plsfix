// Chapter 4, second half: the chart builders, consistent rounding with the two
// custom functions, and unpivot. Owns the sections that write something new
// into the workbook rather than reformatting what is already there.

use super::{Block, Section, section};

pub fn sections() -> Vec<Section> {
    vec![
        waterfall(),
        tornado(),
        football_field(),
        comps_stats(),
        chart_format(),
        rounding(),
        unpivot(),
    ]
}

fn waterfall() -> Section {
    section(
        "Waterfall from selection",
        vec![
            Block::Para(
                r#"Poga "Waterfall from selection" izveido tiltu diagrammu no divu kolonnu tabulas: pirmajā kolonnā nosaukumi, otrajā skaitļi. Pirmā rinda ir sākuma kopsumma, pēdējā rinda ir beigu kopsumma, un starp tām ir soļi."#,
            ),
            Block::Steps(&[
                "Atlasiet abas kolonnas kopā ar pirmo un pēdējo rindu.",
                r#"Nospiediet "Waterfall from selection"."#,
                "Diagramma parādās blakus atlasītajam blokam, nepārklājot citas diagrammas.",
                "Ar peles labo pogu noklikšķiniet uz beigu kolonnas un izvēlieties Set as Total.",
            ]),
            Block::Para(
                r#"Pēdējais solis jāveic ar roku: Office.js šo diagrammas īpašību neatļauj uzstādīt no programmas. Krāsas sākuma, beigu, augšupejošajām un lejupejošajām kolonnām nāk no jūsu paletes."#,
            ),
        ],
    )
}

fn tornado() -> Section {
    section(
        "Tornado from selection",
        vec![
            Block::Para(
                r#"Poga "Tornado from selection" izveido jutīguma diagrammu no trīs kolonnu bloka: faktors, zemākā vērtība un augstākā vērtība."#,
            ),
            Block::Para(
                r#"Bāzes vērtība ir skaitlis virs zemāko un augstāko vērtību kolonnām. Ja tāda nav, par bāzi tiek ņemts abu kolonnu vidējais. Faktori tiek sakārtoti pēc svārstības lieluma, un starpības tiek ierakstītas palīgblokā pa labi no atlases."#,
            ),
            Block::Para(
                r#"Stabiņu galos ir tikai vērtības, tajā pašā skaitļu formātā, kādā atlasē ir zemākā un augstākā vērtība."#,
            ),
        ],
    )
}

fn football_field() -> Section {
    section(
        "Football field",
        vec![
            Block::Para(
                r#"Poga "Football field" no trīs kolonnu tabulas (metode, zemākā vērtība, augstākā vērtība) uzzīmē vērtēšanas diapazonu diagrammu: katrai metodei viena peldoša josla no zemākās līdz augstākajai vērtībai, pirmā rinda augšā. Diagramma ir Excel joslu diagramma ar neredzamu pamata sēriju, tāpēc to var rediģēt kā jebkuru citu diagrammu."#,
            ),
            Block::Steps(&[
                "Atlasiet trīs kolonnas: nosaukumi, zemākā un augstākā vērtība, no 2 līdz 20 rindām. Virsrakstu rinda tiek izlaista, ceturtā un tālākās kolonnas netiek izmantotas.",
                "Nospiediet \"Football field\". Blakus atlasei tiek ierakstīts palīgbloks (metode, zemākā vērtība, diapazons), un diagramma tiek novietota brīvā vietā pie tā.",
                "Ja kādā rindā zemākā vērtība ir lielāka par augstāko, rīks tās apmaina un paziņo, cik rindu apmainīts.",
            ]),
            Block::Para(
                r#"Vērtību ass pārņem skaitļa formātu no zemākās vērtības kolonnas, leģenda ir izslēgta, diagrammas nosaukums ir "Valuation range"."#,
            ),
        ],
    )
}

fn comps_stats() -> Section {
    section(
        "Comps stats",
        vec![
            Block::Para(
                r#"Poga "Comps stats" zem atlasītās salīdzināmo uzņēmumu tabulas ieraksta sešas rindas ar dzīvām formulām: minimums, 25. procentile, mediāna, vidējais, 75. procentile un maksimums. Katra skaitļu kolonna saņem savas formulas, pirmā kolonna nes rindu nosaukumus, un katra šūna pārņem skaitļa formātu no kolonnas pēdējās datu rindas."#,
            ),
            Block::Steps(&[
                "Atlasiet tabulu kopā ar virsrakstu rindu, ja tāda ir: rinda, kurā ir tikai teksts, tiek uzskatīta par virsrakstu un statistikā neietilpst.",
                "Nospiediet \"Comps stats\". Bloks tiek ierakstīts vienu tukšu rindu zem tabulas un paliek atlasīts.",
                "Mainot kādu reizinātāju tabulā, statistika pārrēķinās pati, jo šūnās ir formulas, nevis vērtības.",
            ]),
            Block::Para(
                r#"Rīks atsakās strādāt, ja zem tabulas nav sešu tukšu rindu, ja atlasīti vairāki apgabali, ja datu rindu ir mazāk par divām vai ja nevienā kolonnā nav skaitļu. Ierakstīto bloku var atsaukt ar "Undo last pls,fix action"."#,
            ),
        ],
    )
}

fn chart_format() -> Section {
    section(
        "Brand-format chart un CAGR label",
        vec![
            Block::Para(
                r#"Poga "Brand-format chart" pārnoformē atlasīto diagrammu: fonts, sēriju krāsas un leģenda nāk no jūsu paletes, palīglīnijas tiek izslēgtas, un datu etiķetēs paliek tikai vērtības (bez kategoriju nosaukumiem, sēriju nosaukumiem un procentiem). Sektoru diagrammai etiķetes tiek novietotas ārpus sektoriem, un kategoriju nosaukumus rāda leģenda. Vispirms jāatlasa diagramma, citādi panelis paziņo "Select a chart first"."#,
            ),
            Block::Para(
                r#"Poga "CAGR label" pieraksta atlasītās sērijas vidējo gada pieaugumu diagrammas augšējā labajā stūrī."#,
            ),
        ],
    )
}

fn rounding() -> Section {
    section(
        "Consistent rounding un funkcijas PLSFIX",
        vec![
            Block::Para(
                r#"Noapaļojot katru skaitli atsevišķi, daļu summa bieži neatbilst noapaļotajai kopsummai. Poga "Consistent rounding" šo problēmu novērš: tā ieraksta blakus atlasītajai rindai vai kolonnai formulas, kuru rezultāti summējas tieši ar noapaļoto kopsummu."#,
            ),
            Block::Para(
                r#"Formulas izmanto divas pievienojumprogrammas funkcijas, kuras var rakstīt arī pašrocīgi."#,
            ),
            Block::Table {
                head: &["Funkcija", "Ko tā atgriež"],
                rows: &[
                    &[
                        "=PLSFIX.ROUND(apgabals; numurs; zīmes)",
                        "vienas šūnas daļu no kopīgi noapaļotas grupas",
                    ],
                    &[
                        "=PLSFIX.ROUNDSUM(apgabals; zīmes)",
                        "grupas kopsummu, noapaļotu tā, lai tā sakristu ar daļu summu",
                    ],
                    &[
                        "=PLSFIX.CAGR(sākums; beigas; periodi)",
                        "vidējo gada pieauguma tempu no sākuma vērtības līdz beigu vērtībai norādītajā periodu skaitā",
                    ],
                ],
            },
            Block::Para(
                r#"Argumentu "zīmes" saprot tāpat kā Excel funkcijā ROUND: 0 nozīmē veselus skaitļus, negatīvs skaitlis noapaļo līdz desmitiem vai tūkstošiem."#,
            ),
            Block::Para(
                r#"Funkcija "=PLSFIX.CAGR(100; 200; 4)" atgriež 0,1892 jeb 18,92 % gadā. Negatīva vai nulles vērtība un periodu skaits, kas mazāks par 1, dod #VALUE!."#,
            ),
            Block::Para(
                r#"Šīs funkcijas nāk no pievienojumprogrammas. Datorā, kur pls,fix nav uzstādīts, šādas šūnas rāda #NAME?."#,
            ),
        ],
    )
}

fn unpivot() -> Section {
    section(
        "Unpivot selection",
        vec![
            Block::Para(
                r#"Poga "Unpivot selection" pārraksta krusttabulu par garu sarakstu. Virsrakstu rinda un pirmā kolonna kļūst par kolonnām Row un Column, bet katra šūna par vienu rindu kolonnā Value."#,
            ),
            Block::Para(
                r#"Rezultāts tiek ierakstīts jaunā lapā, tukšās šūnas tiek izlaistas, un sākotnējie dati paliek neskarti."#,
            ),
        ],
    )
}
