// Chapter 2: how the add-in reaches a user, centrally or by hand, what the
// ribbon looks like afterwards, and the demo files a new user can try it on.
// Owns the install steps; the ribbon labels come from manifest.prod.xml and
// must match it word for word.

use super::{Block, Chapter, Section, section};

pub fn setup() -> Chapter {
    Chapter {
        title: "Uzstādīšana",
        sections: vec![central(), manual_install(), ribbon(), demo_files()],
    }
}

fn central() -> Section {
    section(
        "Centralizēta izplatīšana",
        vec![
            Block::Para(
                r#"Parastais ceļš ir centralizēta izplatīšana, ko veic Microsoft 365 administrators. Lietotājam pašam nav jādara nekas: pievienojumprogramma parādās Excel un PowerPoint lentē pati."#,
            ),
            Block::Steps(&[
                "Administrators atver Microsoft 365 administrācijas centru.",
                r#"Sadaļā "Settings" izvēlas "Integrated apps" un tad "Upload custom apps"."#,
                "Augšupielādē manifesta failu manifest.prod.xml.",
                "Piešķir pievienojumprogrammu lietotājiem vai grupai, kurai tā nepieciešama.",
                "Lietotāji restartē Excel un PowerPoint. Cilne parādās, kad Office ir saņēmis piešķīrumu.",
            ]),
            Block::Para(
                r#"Piekļuvi nosaka tieši šis grupas piešķīrums, tāpēc pievienojumprogrammu redz tikai tie, kam tā ir piešķirta."#,
            ),
            Block::Para(
                r#"Pēc izplatīšanas programmas kods tiek atjaunināts serverī, un lietotājam nekas nav jādara. Atkārtota augšupielāde nepieciešama tikai tad, kad mainās pats manifests, piemēram, kad lentei tiek pievienota jauna poga. Faili, kas publicē funkcijas =PLSFIX.ROUND un =PLSFIX.ROUNDSUM, Office kešatmiņā glabājas atsevišķi, un pēc Microsoft dokumentācijas to atjauninājums pie lietotājiem var nonākt 24 stundu laikā."#,
            ),
        ],
    )
}

fn manual_install() -> Section {
    section(
        "Manuāla uzstādīšana",
        vec![
            Block::Para(
                r#"Manuālā uzstādīšana der pārbaudei vai vienam datoram, kamēr centralizētā izplatīšana vēl nav sagatavota. Nepieciešams manifesta fails manifest.prod.xml."#,
            ),
            Block::Para("Mac datorā:"),
            Block::Steps(&[
                "Aizveriet Excel un PowerPoint.",
                "Iekopējiet manifesta failu mapē ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef un, ja nepieciešams PowerPoint, arī mapē ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef.",
                "Atveriet programmu no jauna. Lentē parādās cilne \"pls,fix\".",
            ]),
            Block::Para("Windows datorā:"),
            Block::Steps(&[
                "Novietojiet manifesta failu koplietotā tīkla mapē.",
                r#"Excel izvēlnē "File" atveriet "Options", tad "Trust Center", tad "Trust Center Settings" un "Trusted Add-in Catalogs"."#,
                r#"Pievienojiet mapes adresi, atzīmējiet "Show in Menu" un saglabājiet."#,
                r#"Restartējiet Excel un izvēlnē "Insert" atveriet "My Add-ins", tad cilni "Shared Folder" un izvēlieties pls,fix."#,
            ]),
            Block::Para("Pārlūkprogrammā (Office tīmeklī):"),
            Block::Steps(&[
                r#"Atveriet darbgrāmatu un izvēlnē "Insert" izvēlieties "Add-ins"."#,
                r#"Nospiediet "Upload My Add-in" un norādiet manifesta failu."#,
                "Panelis ir pieejams uzreiz, bet tikai šajā pārlūkprogrammā.",
            ]),
        ],
    )
}

fn ribbon() -> Section {
    section(
        "Lente un paneļa atvēršana",
        vec![
            Block::Para(
                r#"Pēc uzstādīšanas Excel un PowerPoint lentē ir cilne "pls,fix". Excel pusē tajā ir piecas grupas, PowerPoint pusē viena."#,
            ),
            Block::Table {
                head: &["Grupa Excel lentē", "Pogas"],
                rows: &[
                    &[
                        r#""Model Tools""#,
                        r#""Model Tools", "Autocolor", "Fill Right", "Fill Down", "IFERROR""#,
                    ],
                    &[
                        r#""Audit""#,
                        r#""Audit overlay", "Precedents", "Dependents""#,
                    ],
                    &[
                        r#""Paste""#,
                        r#""pls,fix Undo", "Paste values", "Paste formats""#,
                    ],
                    &[r#""Model""#, r#""CAGR", "Sign flip", "x1000", "/1000""#],
                    &[r#""Workbook""#, r#""Waterfall", "Contents sheet""#],
                ],
            },
            Block::Para(
                r#"Katrai lentes pogai ir sava ikona, tāpēc darbības var atpazīt bez uzraksta."#,
            ),
            Block::Para(
                r#"Paneli atver poga "Model Tools" grupā "Model Tools" vai taustiņu kombinācija Ctrl+Shift+M. PowerPoint pusē paneli atver poga "Links" grupā "Links"."#,
            ),
            Block::Para(
                r#"Lentes pogas un taustiņu kombinācijas strādā arī tad, kad panelis nav atvērts."#,
            ),
        ],
    )
}

fn demo_files() -> Section {
    section(
        "Paraugdarbgrāmata un paraugprezentācija",
        vec![
            Block::Para(
                r#"Katram izlaidumam GitHub lapā ir pievienota paraugdarbgrāmata pls,fix Demo Model.xlsx un paraugprezentācija pls,fix Demo Deck.pptx. Abos failos var izmēģināt pievienojumprogrammu, pirms sākat darbu ar savu darbgrāmatu."#,
            ),
            Block::Para(
                r#"Paraugdarbgrāmatā katras redzamās lapas augšā ir īss uzdevumu saraksts "Try on this sheet" un rinda "Commands" ar vajadzīgajām lentes pogām un taustiņu kombinācijām."#,
            ),
            Block::Para(
                r#"Paraugprezentācijā ir četri uzdevumi: katrā jāievieto kāds objekts izvēlētajā slaidā un vietā, izmantojot izvēlnes "Slide" un "Where" (skatiet nodaļu "Saites uz PowerPoint", sadaļu "Slaids un vieta")."#,
            ),
        ],
    )
}
