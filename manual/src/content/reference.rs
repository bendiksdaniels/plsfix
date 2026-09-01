// Chapters 8, 9 and 10: the shortcut list, the troubleshooting tables and the
// version note. Owns the Latvian gloss for every shortcut action id; the keys
// and English names themselves come from public/shortcuts.json at build time.

use super::{Block, Chapter, Section, section};

pub fn shortcuts() -> Chapter {
    Chapter {
        title: "Īsinājumtaustiņi",
        sections: vec![section(
            "Visas kombinācijas",
            vec![
                Block::Para(
                    r#"Visas kombinācijas strādā arī tad, kad panelis nav atvērts. Ja kombinācija pārklājas ar Excel savējo, Excel vienreiz jautā, kuru darbību paturēt. Skaitļu formātu cikli apzināti atrodas uz tiem pašiem taustiņiem, ko Excel lieto formātiem, lai ierastā kustība nonāktu pie jūsu paletes formāta."#,
                ),
                Block::Para(
                    r#"Kombinācijas var mainīt Office pievienojumprogrammu īsinājumtaustiņu iestatījumos."#,
                ),
                Block::Shortcuts,
            ],
        )],
    }
}

/// The Latvian gloss for one shortcut action id, empty when the id is unknown.
pub fn gloss(id: &str) -> &'static str {
    GLOSS
        .iter()
        .find(|(key, _)| *key == id)
        .map(|(_, text)| *text)
        .unwrap_or("")
}

const GLOSS: &[(&str, &str)] = &[
    ("PLSFIX_SHOWPANE", "Atver pls,fix paneli"),
    ("PLSFIX_AUTOCOLOR", "Nokrāso atlasi pēc satura"),
    ("PLSFIX_FILLRIGHT", "Izplata formulu pa labi"),
    ("PLSFIX_FILLDOWN", "Izplata formulu uz leju"),
    ("PLSFIX_IFERROR", "Uzliek vai noņem IFERROR aizsargu"),
    ("PLSFIX_SCALEUP", "Reizina atlasi ar tūkstoti"),
    ("PLSFIX_SCALEDOWN", "Dala atlasi ar tūkstoti"),
    ("PLSFIX_CYC_GENERAL", "Cikls: parastie skaitļi"),
    ("PLSFIX_CYC_DATE", "Cikls: datumi"),
    ("PLSFIX_CYC_CURRENCY", "Cikls: valūta"),
    ("PLSFIX_CYC_PERCENT", "Cikls: procenti"),
    ("PLSFIX_CYC_MULTIPLE", "Cikls: reizinājumi"),
    ("PLSFIX_CYC_TITLE", "Cikls: virsraksta rindas stils"),
    ("PLSFIX_CYC_RESULT", "Cikls: rezultāta rindas stils"),
    ("PLSFIX_CYC_ITEM", "Cikls: posteņa rindas stils"),
    ("PLSFIX_CYC_FILL", "Cikls: aizpildījuma krāsa"),
    ("PLSFIX_CYC_FONT", "Cikls: fonta krāsa"),
    ("PLSFIX_CYC_BORDER", "Cikls: apmales"),
    ("PLSFIX_CYC_ROWH", "Cikls: rindas augstums"),
    ("PLSFIX_CYC_COLW", "Cikls: kolonnas platums"),
    ("PLSFIX_PAINT_APP1", "Uzliek pirmo otas slotu"),
    ("PLSFIX_PAINT_APP2", "Uzliek otro otas slotu"),
    ("PLSFIX_PAINT_APP3", "Uzliek trešo otas slotu"),
    ("PLSFIX_AUDIT", "Ieslēdz vai izslēdz pārbaudes slāni"),
    (
        "PLSFIX_TRACE_PRE",
        "Atlasa šūnas, no kurām lasa aktīvā šūna",
    ),
    (
        "PLSFIX_TRACE_DEP",
        "Atlasa šūnas, kas lasa no aktīvās šūnas",
    ),
    ("PLSFIX_UNDO", "Atceļ pēdējo pls,fix darbību"),
    ("PLSFIX_COPYSRC", "Atzīmē atlasi kā kopēšanas avotu"),
    ("PLSFIX_PASTE_VALUES", "Ielīmē avota vērtības"),
    ("PLSFIX_PASTE_FORMATS", "Ielīmē avota noformējumu"),
    ("PLSFIX_PASTE_EXACT", "Ielīmē formulas, atsauces nemainot"),
    ("PLSFIX_PASTE_TRANSPOSE", "Ielīmē rindas kā kolonnas"),
    ("PLSFIX_CAGR", "Ieraksta CAGR formulu blakus atlasei"),
    ("PLSFIX_SIGN", "Maina atlases zīmi uz pretējo"),
    ("PLSFIX_DEC_MORE", "Pievieno vienu zīmi aiz komata"),
    ("PLSFIX_DEC_LESS", "Noņem vienu zīmi aiz komata"),
    ("PLSFIX_WATERFALL", "Izveido tiltu diagrammu no atlases"),
    ("PLSFIX_CHARTFMT", "Pārnoformē diagrammu pēc paletes"),
    (
        "PLSFIX_CHART_CAGR",
        "Pieraksta pieauguma tempu uz diagrammas",
    ),
    ("PLSFIX_TOC", "Izveido satura rādītāja lapu"),
    ("PLSFIX_FIND", "Meklē visā darbgrāmatā"),
    ("PLSFIX_STYLES_SCAN", "Atrod neizmantotos šūnu stilus"),
];

pub fn problems() -> Chapter {
    Chapter {
        title: "Biežākās problēmas",
        sections: vec![messages(), limits()],
    }
}

fn messages() -> Section {
    section(
        "Ziņojumi un to iemesli",
        vec![Block::Table {
            head: &["Ziņojums vai pazīme", "Cēlonis", "Ko darīt"],
            rows: MESSAGES,
        }],
    )
}

const MESSAGES: &[&[&str]] = &[
    &[
        r#"Panelis paliek uz "Connecting""#,
        "Office nav paziņojis paneli par gatavu. Visbiežāk tas notiek pārlūkprogrammā, kad neizdodas startēt funkciju izpildvidi.",
        "Pēc dažām sekundēm panelis startē ierobežotā režīmā un par to paziņo. Aizveriet un atveriet darbgrāmatu no jauna vai pārlādējiet lapu.",
    ],
    &[
        r#""Select a chart first, or pick one from the list.""#,
        "Diagramma nav atlasīta, un lapā ir vairākas diagrammas.",
        "Atlasiet diagrammu vai izvēlieties to no saraksta zem pogas.",
    ],
    &[
        r#""No chart on this sheet.""#,
        "Aktīvajā lapā nav nevienas diagrammas.",
        "Pārejiet uz lapu, kurā diagramma ir.",
    ],
    &[
        r#""source missing""#,
        "Excel pusē avota nosaukums vai diagramma ir dzēsta.",
        r#"Excel cilnē "Links" pārbaudiet rindu ar "Go to source" un, ja nepieciešams, eksportējiet saiti no jauna."#,
    ],
    &[
        "Šūnas ar =PLSFIX.ROUND rāda #NAME?",
        "Excel vēl nav saņēmis pievienojumprogrammas funkcijas.",
        "Aizveriet un atveriet darbgrāmatu no jauna. Datorā, kur pls,fix nav uzstādīts, šīs funkcijas nedarbosies.",
    ],
    &[
        "Excel rāda citus atdalītājus, nekā paredz izvēlētā valoda",
        "Atdalītāji ir Excel iestatījums, ko panelis var tikai nolasīt.",
        "Mainiet tos Mac datorā: Excel > Preferences > Edit. Windows datorā: File > Options > Advanced.",
    ],
    &[
        r#""select a single range""#,
        "Atlasē ir vairāki nesaistīti apgabali.",
        r#"Atlasiet vienu nepārtrauktu apgabalu un nospiediet "Export selection" vēlreiz."#,
    ],
    &[
        r#""Select a slide first.""#,
        "PowerPoint slaidu rādītājā nav atlasīts neviens slaids.",
        "Atlasiet slaidu un mēģiniet vēlreiz.",
    ],
    &[
        r#""Wrong link key""#,
        "Attēlā saglabātais saites marķieris neatbilst relejā reģistrētajam.",
        r#"Eksportējiet saiti no jauna un ievietojiet attēlu vēlreiz no "Inbox"."#,
    ],
    &[
        r#"Cilne "Inbox" ir tukša"#,
        "PowerPoint nav ielīmēta atslēga vai ielīmēta ir vecā atslēga.",
        r#"Nokopējiet atslēgu Excel cilnē "Links" un ielīmējiet to PowerPoint cilnē "Settings"."#,
    ],
    &[
        r#""No combination reaches the target within the tolerance.""#,
        "Neviena skaitļu kombinācija atlasē nesasniedz mērķi pieļaujamās novirzes robežās.",
        "Palieliniet toleranci vai pārbaudiet, vai atlasē ir īstie skaitļi.",
    ],
];

fn limits() -> Section {
    section(
        "Ierobežojumi",
        vec![Block::Bullets(&[
            r#"Excel taustiņu kombinācija Ctrl+Z pievienojumprogrammas izmaiņas neatceļ; lietojiet "Undo last pls,fix action"."#,
            "Atcelšana aptver līdz 5 000 šūnām. Rindu augstumi, kolonnu platumi, diagrammas, formas, lapas un definētie nosaukumi paliek ārpus tās.",
            "Rindu stilu cikli vienā reizē apstrādā līdz 500 rindām.",
            "Tiltu diagrammas beigu kolonna jāatzīmē ar roku: Set as Total.",
            r#""Precedents" prasa ExcelApi 1.12, "Dependents" prasa ExcelApi 1.13."#,
            "Pārbaude un izsekošana strādā vienas atvērtās darbgrāmatas robežās.",
            r#"Poga "Prepare for sharing" tālummaiņu atiestatīt nevar, jo Office.js to nepiedāvā."#,
            "PowerPoint saites grupās tiek atrastas trīs līmeņus dziļi un tikai ar PowerPointApi 1.8.",
            r#"Relejs glabā vienu iepriekšējo attēla versiju, tāpēc "Revert last update" atgriež tikai vienu soli."#,
            r#""Find a combination" aptver līdz 34 skaitliskām šūnām vienā atlasē."#,
            r#""Object tools" un "Smart Painter" prasa PowerPoint 2021 vai Microsoft 365 (PowerPointApi 1.5)."#,
            r#""Align" prasa vismaz divas atlasītas figūras, "Distribute" vismaz trīs, "Swap" tieši divas."#,
        ])],
    )
}

pub fn version() -> Chapter {
    Chapter {
        title: "Versija",
        sections: vec![section(
            "Kura versija ir aprakstīta",
            vec![
                Block::Para(
                    "Šī rokasgrāmata apraksta pls,fix versiju {version}. Tā pati versija ir uzrakstīta uz vāka un katras lapas apakšā.",
                ),
                Block::Para(
                    "Uzstādīto versiju var pārbaudīt paneļa apakšējā labajā stūrī. Numuru veido trīs daļas, un pēdējā vienmēr ir trīs ciparu.",
                ),
                Block::Para(
                    "Programmas kods tiek atjaunināts serverī, tāpēc jauna versija nonāk pie lietotāja bez darbībām no viņa puses. Ja mainās manifests, administratoram tas jāaugšupielādē no jauna, un funkciju faili Office kešatmiņā var atjaunoties līdz 24 stundu laikā.",
                ),
            ],
        )],
    }
}
