// Chapters 1 and 3: what pls,fix is, where it runs, and how the task pane is
// laid out. Owns the reader's first two pages; the invariant is that every
// claim here comes from README.md, the manifest or the pane markup.

use super::{Block, Chapter, Section, section};

pub fn what_it_is() -> Chapter {
    Chapter {
        title: "Kas ir pls,fix",
        sections: vec![purpose(), where_it_runs(), what_stays()],
    }
}

pub fn pane() -> Chapter {
    Chapter {
        title: "Panelis un cilnes",
        sections: vec![pane_layout(), pane_tabs(), pane_behaviour()],
    }
}

fn purpose() -> Section {
    section(
        "Kam programma paredzēta",
        vec![
            Block::Para(
                r#"pls,fix ir Microsoft Excel un Microsoft PowerPoint pievienojumprogramma finanšu modelēšanai. Tā paātrina darbu, ko modelētājs atkārto katru dienu: vienāda noformējuma uzlikšanu, šūnu krāsošanu pēc satura, formulu pārbaudi, diagrammu izveidi un tabulu nogādāšanu prezentācijā."#,
            ),
            Block::Para(
                r#"Nosaukums nāk no vēstules, ko pazīst katrs analītiķis. Komats ir daļa no nosaukuma, un tas pats komats ir programmas logotips."#,
            ),
            Block::Para(
                r#"Saskarne ir angļu valodā. Rokasgrāmatā pogu un ciļņu nosaukumi ir doti tieši tā, kā tie redzami ekrānā, un likti pēdiņās, piemēram "Autocolor selection"."#,
            ),
        ],
    )
}

fn where_it_runs() -> Section {
    section(
        "Kur tā darbojas",
        vec![
            Block::Para(
                r#"pls,fix ir tīmekļa pievienojumprogramma. Tā neprasa instalēšanu datorā, un viena un tā pati saskarne strādā visās vidēs, kur Office ir pieejams."#,
            ),
            Block::Bullets(&[
                r#"Excel ar Microsoft 365 abonementu: Windows datorā, Mac datorā un pārlūkprogrammā."#,
                r#"PowerPoint ar Microsoft 365 abonementu: Windows datorā, Mac datorā un pārlūkprogrammā."#,
                r#"Saitēm starp Excel un PowerPoint nepieciešams PowerPoint 2504 vai jaunāka versija Windows datorā un PowerPoint 16.96 vai jaunāka versija Mac datorā."#,
                r#"Panelis ielādējas no adreses dbautomatizacijas.com/modelis/, tāpēc darba laikā nepieciešams interneta pieslēgums."#,
            ]),
            Block::Para(
                r#"Divas darbības ir atkarīgas no Excel versijas: "Precedents" prasa ExcelApi 1.12, "Dependents" prasa ExcelApi 1.13. Vecākā Excel versijā panelis par to pasaka tieši, bet pārējās darbības strādā."#,
            ),
        ],
    )
}

fn what_stays() -> Section {
    section(
        "Kas paliek jūsu datorā",
        vec![
            Block::Para(
                r#"Darbgrāmatas dati paliek Excel procesā. Panelis tos nolasa un ieraksta atpakaļ, bet nekur nesūta."#,
            ),
            Block::Para(
                r#"Vienīgais, kas atstāj datoru, ir uz PowerPoint saistītā objekta attēls. Tas tiek šifrēts panelī pirms augšupielādes, un relejs glabā tikai šifrētus datus, nevis failu nosaukumus vai skaitļus. Darbgrāmata bez saitēm nesūta neko."#,
            ),
            Block::Para(
                r#"Arī logotipa krāsu nolasīšana notiek panelī: augšupielādētais attēls nekur netiek nosūtīts."#,
            ),
        ],
    )
}

fn pane_layout() -> Section {
    section(
        "Paneļa uzbūve",
        vec![
            Block::Para(
                r#"Panelis atveras Excel loga labajā pusē un paliek atvērts, kamēr to aizver. Augšpusē ir komata logotips, nosaukums pls,fix un virsraksts "Build cleaner models."."#,
            ),
            Block::Para(
                r#"Zem virsraksta ir stāvokļa rinda. Kamēr panelis pieslēdzas, tajā redzams "Connecting". Kad savienojums ir izveidots, tajā redzams "Excel connected", bet PowerPoint pusē "PowerPoint connected"."#,
            ),
            Block::Para(
                r#"Paneļa apakšā ir rinda "Workbook data stays inside Excel" un pašreizējā versija. PowerPoint pusē tajā pašā vietā ir rinda "Linked objects stay inside PowerPoint"."#,
            ),
            Block::Image {
                file: "excel-brand-language.png",
                alt: r#"pls,fix panelis Excel loga labajā pusē: komata logotips, stāvokļa rinda "Excel connected" un ciļņu josla."#,
            },
        ],
    )
}

fn pane_tabs() -> Section {
    section(
        "Četras cilnes",
        vec![
            Block::Para(r#"Zem stāvokļa rindas ir ciļņu josla. Excel panelī ir četras cilnes."#),
            Block::Table {
                head: &["Cilne", "Ko tā satur"],
                rows: &[
                    &[
                        r#""Tools""#,
                        "Darbs ar atlasi: stili, skaitļu formāti, aizpildīšana, ielīmēšana, krāsošana, formulu pārbaude un diagrammas.",
                    ],
                    &[
                        r#""Workbook""#,
                        "Darbs ar visu darbgrāmatu: lapu saraksts, meklēšana, satura rādītājs, nosaukumu un stilu tīrīšana, sagatavošana nosūtīšanai.",
                    ],
                    &[
                        r#""Links""#,
                        "Saites uz PowerPoint: eksports, saišu saraksts, saites atslēga.",
                    ],
                    &[
                        r#""Brand""#,
                        "Uzņēmuma krāsas, fonts, valoda, valūta un iestatījumu imports vai eksports.",
                    ],
                ],
            },
            Block::Para(
                r#"PowerPoint panelī ir trīs cilnes: "Links", "Inbox" un "Settings". Tās aprakstītas nodaļā par saitēm."#,
            ),
        ],
    )
}

fn pane_behaviour() -> Section {
    section(
        "Kā panelis strādā ar atlasi",
        vec![
            Block::Para(
                r#"Gandrīz katra darbība attiecas uz to, kas Excel ir atlasīts tajā brīdī. Vispirms atlasiet šūnas vai diagrammu, tad nospiediet pogu."#,
            ),
            Block::Para(
                r#"Pēc katras darbības paneļa apakšā uz brīdi parādās paziņojums ar rezultātu vai ar kļūdas iemeslu. Ja atlase darbībai neder, paziņojums pasaka, kas nepieciešams."#,
            ),
            Block::Para(
                r#"Excel taustiņu kombinācija Ctrl+Z neatceļ pievienojumprogrammas veiktās izmaiņas, jo Office.js ierakstus Excel savā atcelšanas sarakstā neieliek. Šim nolūkam ir atsevišķa poga "Undo last pls,fix action"."#,
            ),
        ],
    )
}
