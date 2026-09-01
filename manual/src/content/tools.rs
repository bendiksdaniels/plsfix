// Chapter 4, first half: the "Tools" tab from the selection inspector down to
// the trace panel. Owns the everyday formatting and editing sections; the
// chart, rounding and unpivot sections live in `charts.rs`.

use super::charts;
use super::{Block, Chapter, Section, section};

pub fn tools() -> Chapter {
    let mut sections = vec![
        overview(),
        inspector(),
        reconcile(),
        formatting(),
        paintbrush(),
        cycles(),
        fill_and_paste(),
        iferror(),
        autocolor(),
        audit(),
        trace(),
        quick_maths(),
    ];
    sections.extend(charts::sections());
    sections.push(undo());
    Chapter {
        title: "Cilne Tools",
        sections,
    }
}

fn overview() -> Section {
    section(
        "Kas atrodas cilnē",
        vec![
            Block::Para(
                r#"Cilne "Tools" ir paneļa darba daļa. Tā ir sadalīta blokos: atlases pārskats, veidnes, noformējums, formātu cikli, modeļa rīki, pārbaude un diagrammas."#,
            ),
            Block::Image {
                file: "excel-tools.png",
                alt: r#"Cilne "Tools": atlases pārskats un noformējuma pogas "Title", "Header", "Input", "Formula", "Result" un "Clear"."#,
            },
        ],
    )
}

fn inspector() -> Section {
    section(
        "Atlases pārskats",
        vec![
            Block::Para(
                r#"Bloks "Selection inspector" rāda atlasītā apgabala adresi un četrus skaitļus: cik šūnu ("cells"), cik formulu ("formulas"), cik kļūdu ("errors") un cik tukšu šūnu ("blank")."#,
            ),
            Block::Para(
                r#"Tas ir ātrākais veids, kā pārbaudīt, vai atlasē nav palikusi kļūda vai tukša vieta. Poga ar apļveida bultiņu blakus virsrakstam pārlasa atlasi."#,
            ),
        ],
    )
}

fn reconcile() -> Section {
    section(
        "Atrast kombināciju",
        vec![
            Block::Para(
                r#"Sadaļa "Find a combination" atrod, kuras atlasītā apgabala šūnas kopā dod vēlamo summu. Tas noder, piemēram, pārbaudot, kuri posteņi kopā veido starpību starp divām kopsummām."#,
            ),
            Block::Steps(&[
                "Atlasiet vienu nepārtrauktu apgabalu ar skaitļiem.",
                r#"Laukā "Target" ierakstiet meklējamo summu, laukā "Tolerance" pieļaujamo novirzi."#,
                r#"Nospiediet "Find cells"."#,
            ]),
            Block::Para(
                r#"Apgabalā drīkst būt līdz 34 skaitliskām šūnām; teksts, tukšas šūnas un formulu kļūdas netiek skaitītas. Ja atlasē ir vairāk, panelis lūdz atlasīt mazāku apgabalu."#,
            ),
            Block::Para(
                r#"Kad kombinācija atrasta, tieši tās šūnas kļūst par jauno atlasi. Zem pogas parādās, cik šūnu atrasts, to summa ("Sum") un atlikusī novirze no mērķa ("Variance"). Ja neviena kombinācija tolerances robežās neatbilst, panelis to pasaka."#,
            ),
        ],
    )
}

fn formatting() -> Section {
    section(
        "Noformējuma stili un skaitļu formāti",
        vec![
            Block::Para(
                r#"Bloks "Model formatting" uzliek atlasei vienu no pieciem modeļa stiliem: "Title", "Header", "Input", "Formula" un "Result". Poga "Clear" noņem noformējumu."#,
            ),
            Block::Para(
                r#"Zem stiliem ir četras skaitļu formātu pogas: "1,234" veseliem skaitļiem, "1,234.0" ar vienu zīmi aiz komata, valūtas poga un "12.3%" procentiem. Valūtas pogas uzraksts seko iestatītajai valūtai un valodai cilnē "Brand"."#,
            ),
            Block::Para(
                r#"Visi stili ņem krāsas un fontu no cilnes "Brand", tāpēc viena darbgrāmata izskatās vienādi neatkarīgi no tā, kurš to noformē."#,
            ),
        ],
    )
}

fn paintbrush() -> Section {
    section(
        "Otas sloti",
        vec![
            Block::Para(
                r#"Pogas "Save 1", "Save 2" un "Save 3" iegaumē aktīvās šūnas skaitļu formātu, fontu, aizpildījumu, līdzinājumu un apmales. Pogas "Use 1", "Use 2" un "Use 3" uzliek iegaumēto noformējumu atlasei."#,
            ),
            Block::Para(
                r#"Sloti tiek saglabāti pašā darbgrāmatā, tāpēc kolēģis, kas atver to pašu failu citā datorā, redz tos pašus trīs slotus. Ja darbgrāmata vēl nekad nav saglabājusi nevienu slotu, tiek izmantoti šī datora sloti. Zem pogām redzams, kas katrā slotā ir iegaumēts."#,
            ),
        ],
    )
}

fn cycles() -> Section {
    section(
        "Formātu cikli",
        vec![
            Block::Para(
                r#"Bloks "Format cycles" strādā citādi nekā parasta poga: nospiežot to pašu pogu vai to pašu taustiņu kombināciju vēlreiz, formāts pāriet uz nākamo variantu tajā pašā saimē."#,
            ),
            Block::Table {
                head: &["Poga", "Varianti"],
                rows: &[
                    &[r#""1,234""#, "1,234 / 1,234.0 / 1,234.00"],
                    &[r#""Date""#, "31.12.2026 / Dec-26 / 2026"],
                    &[
                        r#""Cur.""#,
                        "valūta bez zīmēm aiz komata / ar vienu zīmi / tūkstošos",
                    ],
                    &[r#""%""#, "12.3% / 12% / 12.34%"],
                    &[r#""0.0x""#, "1.5x / 1.50x"],
                    &[
                        r#""Title", "Result", "Item""#,
                        "rindu stili attiecīgajai rindas lomai",
                    ],
                    &[
                        r#""Fill", "Font""#,
                        "aizpildījuma un fonta krāsas no jūsu paletes",
                    ],
                    &[
                        r#""Border""#,
                        "apakšējā līnija / kopsummas / rezultāta / rāmis / režģis",
                    ],
                    &[r#""Row height""#, "15 / 18 / 21 / 24 / 30 pt"],
                    &[r#""Column width""#, "64 / 80 / 96 / 120 / 48 pt"],
                ],
            },
            Block::Para(
                r#"Rindu stilu cikli strādā pa rindām un vienā reizē apstrādā līdz 500 rindām."#,
            ),
        ],
    )
}

fn fill_and_paste() -> Section {
    section(
        "Aizpildīšana un ielīmēšana",
        vec![
            Block::Para(
                r#"Pogas "Fill formula right" un "Fill formula down" izplata pirmās šūnas formulu pa atlasi. Ja atlasīta ir tikai viena šūna, apjomu nosaka blakus esošā rinda vai kolonna, tāpēc atlase iepriekš nav jāvelk."#,
            ),
            Block::Para(
                r#"Ielīmēšanas rinda strādā divos soļos. Vispirms atlasiet avotu un nospiediet "Mark". Pēc tam atlasiet mērķi un izvēlieties vienu no pogām."#,
            ),
            Block::Table {
                head: &["Poga", "Ko tā ielīmē"],
                rows: &[
                    &[r#""Values""#, "tikai vērtības"],
                    &[r#""Formats""#, "tikai noformējumu"],
                    &[
                        r#""Exact""#,
                        "formulas tieši tādas, kādas tās ir, atsauces nemainot",
                    ],
                    &[r#""Transpose""#, "rindas kā kolonnas"],
                ],
            },
        ],
    )
}

fn iferror() -> Section {
    section(
        "IFERROR aizsargs",
        vec![Block::Para(
            r#"Poga "IFERROR guard" apliek atlasītās formulas ar IFERROR(…, 0). Nospiežot to vēlreiz, aizsargs tiek noņemts, tāpēc darbība ir atgriezeniska."#,
        )],
    )
}

fn autocolor() -> Section {
    section(
        "Autocolor",
        vec![
            Block::Para(
                r#"Poga "Autocolor selection" nokrāso atlasītās šūnas pēc to satura. Krāsas nāk no cilnes "Brand"."#,
            ),
            Block::Bullets(&[
                "Ievadītie skaitļi: viena krāsa (pēc noklusējuma zila).",
                "Formulas: otra krāsa (pēc noklusējuma melna).",
                "Atsauces uz citu lapu: trešā krāsa (pēc noklusējuma zaļa).",
                "Atsauces uz citu darbgrāmatu: atsevišķa krāsa.",
                "Formulas, kurās ierakstīts skaitlis: atsevišķa krāsa.",
            ]),
            Block::Para(
                r#"Poga "Insert color key" ieliek lapā leģendu ar krāsu skaidrojumu, sākot no aktīvās šūnas."#,
            ),
            Block::Image {
                file: "excel-autocolor.png",
                alt: "Peļņas un zaudējumu aprēķins pēc krāsošanas: ievadītie skaitļi, formulas un atsauces uz citu lapu katra savā krāsā.",
            },
        ],
    )
}

fn audit() -> Section {
    section(
        "Formulu pārbaude",
        vec![
            Block::Para(
                r#"Poga "Audit overlay" uzliek atlasei pārbaudes slāni. Šūnas, kuru formula sakrīt ar kaimiņu formulu, tiek iekrāsotas svītraini. Šūnas, kas no rindas vai kolonnas parauga atšķiras, tiek iekrāsotas sarkanīgi."#,
            ),
            Block::Para(
                r#"Blakus pogai redzams stāvoklis "Off" vai "On". Nospiežot pogu vēlreiz, slānis tiek noņemts un sākotnējais aizpildījums atjaunots."#,
            ),
            Block::Para(
                r#"Slāņa pieraksts glabājas pašā darbgrāmatā, tāpēc, atverot failu, kas saglabāts pārbaudes laikā, sākotnējās krāsas var atjaunot."#,
            ),
            Block::Image {
                file: "excel-audit-overlay.png",
                alt: "Pārbaudes slānis peļņas un zaudējumu aprēķinā: svītrotas šūnas ar vienādām formulām un sarkanīgas šūnas, kas no parauga atšķiras.",
            },
        ],
    )
}

fn trace() -> Section {
    section(
        "Precedents un Dependents",
        vec![
            Block::Para(
                r#"Poga "Precedents" atlasa šūnas, no kurām aktīvā šūna lasa. Poga "Dependents" atlasa šūnas, kas lasa no aktīvās šūnas."#,
            ),
            Block::Para(
                r#"Atrastās adreses parādās panelī. Nospiežot uz adreses, Excel pārlec uz to un pārbaudi var turpināt tālāk. Bultiņa pa kreisi atgriež uz iepriekšējo šūnu."#,
            ),
            Block::Para(
                r#"Pārbaude strādā vienas atvērtās darbgrāmatas robežās: Office.js citos failos ieskatīties nevar."#,
            ),
        ],
    )
}

fn quick_maths() -> Section {
    section(
        "CAGR, zīme, mērogs un zīmes aiz komata",
        vec![
            Block::Table {
                head: &["Poga", "Ko tā dara"],
                rows: &[
                    &[
                        r#""CAGR""#,
                        "ieraksta vidējā gada pieauguma formulu tieši aiz atlasītās rindas vai kolonnas",
                    ],
                    &[r#""± sign""#, "maina atlases zīmi uz pretējo"],
                    &[r#""× 1,000""#, "reizina atlasi ar tūkstoti"],
                    &[r#""÷ 1,000""#, "dala atlasi ar tūkstoti"],
                    &[r#"".00 +""#, "pievieno vienu zīmi aiz komata"],
                    &[r#"".00 −""#, "noņem vienu zīmi aiz komata"],
                ],
            },
            Block::Para(r#"Mērogošana strādā gan ar ierakstītiem skaitļiem, gan ar formulām."#),
        ],
    )
}

fn undo() -> Section {
    section(
        "Undo last pls,fix action",
        vec![
            Block::Para(
                r#"Poga "Undo last pls,fix action" atjauno to, ko pēdējā pls,fix darbība pārrakstīja: formulas, skaitļu formātus, aizpildījumus, fontus, apmales, līdzinājumu, teksta aplaušanu un atkāpi. Zem pogas redzams, kura darbība tiks atcelta."#,
            ),
            Block::Para(
                r#"Atcelšana aptver līdz 5 000 šūnām. Lielāka darbība tiek vai nu atteikta, vai izpildīta bez atcelšanas iespējas, un paziņojums to pasaka. Ārpus atcelšanas paliek rindu augstumi, kolonnu platumi un darbības ar diagrammām, formām, lapām un definētiem nosaukumiem."#,
            ),
        ],
    )
}
