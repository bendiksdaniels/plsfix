// Chapters 5 and 7: the "Workbook" tab (whole-file tools) and the "Brand" tab
// (palette, font, language, currency). Owns the two tabs that change how a
// whole file behaves rather than one selection.

use super::{Block, Chapter, Section, section};

pub fn workbook() -> Chapter {
    Chapter {
        title: "Cilne Workbook",
        sections: vec![
            explorer(),
            super_find(),
            contents(),
            names(),
            styles(),
            share(),
            model_check(),
        ],
    }
}

pub fn brand() -> Chapter {
    Chapter {
        title: "Cilne Brand",
        sections: vec![
            palette(),
            logo(),
            language(),
            currency(),
            autocolor_on_edit(),
            transfer(),
        ],
    }
}

fn explorer() -> Section {
    section(
        "Lapu saraksts",
        vec![
            Block::Para(
                r#"Bloks "Sheet explorer" rāda visas darbgrāmatas lapas. Nospiežot uz nosaukuma, Excel pārlec uz šo lapu."#,
            ),
            Block::Para(
                r#"Aplītis rindas labajā pusē lapu paslēpj vai parāda. Ļoti paslēptās lapas sarakstā ir redzamas, bet netiek aiztiktas, jo tās var mainīt tikai ārpus Excel saskarnes."#,
            ),
            Block::Image {
                file: "excel-workbook.png",
                alt: r#"Cilne "Workbook": lapu saraksts ar katras lapas redzamības stāvokli."#,
            },
        ],
    )
}

fn super_find() -> Section {
    section(
        "Super Find",
        vec![
            Block::Para(
                r#"Bloks "Super Find" meklē vienā piegājienā pa visu darbgrāmatu, arī paslēptajās lapās. Ierakstiet meklējamo un nospiediet taustiņu Enter vai pogu "Find in workbook"."#,
            ),
            Block::Para("Meklēšana aptver:"),
            Block::Bullets(&[
                "šūnu vērtības,",
                "formulu tekstu,",
                "darbgrāmatas līmeņa definētos nosaukumus,",
                "lapu nosaukumus,",
                "šūnu komentārus kopā ar atbildēm un autoru.",
            ]),
            Block::Para(
                r#"Trīs izvēles rūtiņas maina meklēšanu: "Match case" prasa sakrist reģistram, "Search formulas" ļauj meklēt aiz vērtības esošajā formulā, "Comments" iekļauj komentārus. Komentāru meklēšanai nepieciešams Excel 365; vecākā versijā panelis pasaka, ka komentāri izlaisti."#,
            ),
            Block::Para(
                r#"Atradumi ir sakārtoti darbgrāmatas secībā, un nospiešana uz atraduma pārlec uz šo šūnu."#,
            ),
        ],
    )
}

fn contents() -> Section {
    section(
        "Satura rādītājs",
        vec![Block::Para(
            r#"Poga "Insert contents sheet" izveido lapu ar saitēm uz katru redzamo lapu. Katrā nospiešanas reizē lapa tiek pārrakstīta no jauna, tāpēc pēc lapu pievienošanas vai pārdēvēšanas pietiek nospiest pogu vēlreiz."#,
        )],
    )
}

fn names() -> Section {
    section(
        "Bojātie nosaukumi",
        vec![
            Block::Para(
                r#"Poga "Scan broken names" atrod definētos nosaukumus, kas norāda uz dzēstām šūnām, tas ir, uz #REF!. Šādi nosaukumi paliek failā pēc rindu vai lapu dzēšanas un vēlāk rada kļūdas."#,
            ),
            Block::Para(
                r#"Zem pogas parādās atrasto nosaukumu skaits. Poga "Delete broken names" tos izdzēš."#,
            ),
        ],
    )
}

fn styles() -> Section {
    section(
        "Neizmantotie stili",
        vec![
            Block::Para(
                r#"Poga "Scan styles" atrod pielāgotos šūnu stilus, ko darbgrāmatā nelieto neviena šūna. Uzkrājušies stili palēnina failu un aizpilda Excel stilu sarakstu."#,
            ),
            Block::Para(
                r#"Poga "Delete unused styles" tos izdzēš pēc apstiprinājuma. Ja kāda lapa ir pārāk liela, lai to pārbaudītu, panelis to nosauc un dzēšanu neļauj, tāpēc daļēja atbilde nekad neizdzēš stilu, kas vēl tiek lietots."#,
            ),
        ],
    )
}

fn share() -> Section {
    section(
        "Prepare for sharing",
        vec![
            Block::Para(
                r#"Poga "Prepare for sharing" sagatavo darbgrāmatu nosūtīšanai. Tā katrā redzamajā lapā atgriež kursoru uz šūnu A1 un atstāj darbgrāmatu atvērtu uz pirmās lapas."#,
            ),
            Block::Para("Pēc tam panelis uzskaita to, ko saņēmējs failā vēl atradīs:"),
            Block::Bullets(&[
                "paslēptās lapas,",
                "saites uz citām darbgrāmatām,",
                "nosaukumus, kas norāda uz dzēstām šūnām,",
                "šūnas ar pievienojumprogrammas funkcijām, kuras citā datorā rādīs #NAME?,",
                "ieslēgtu krāsošanu rediģēšanas laikā,",
                "saišu reģistru, ja darbgrāmatā ir saites uz PowerPoint.",
            ]),
            Block::Para(
                r#"Neviens no šiem elementiem netiek dzēsts, un paslēptās lapas netiek aiztiktas: lēmums paliek jums. Tālummaiņu atiestatīt nav iespējams, jo Office.js šo lapas īpašību nepiedāvā."#,
            ),
        ],
    )
}

fn model_check() -> Section {
    section(
        "Model check",
        vec![Block::Para(
            r#""Model check" vienā piegājienā uzskaita, ko pārbaudītājs modelī atzīmētu: formulu kļūdas, skaitļus formulās, no rindas atšķirīgas formulas, mainīgās funkcijas (OFFSET, INDIRECT, NOW u. c.), sabojātos nosaukumus, nelietotos stilus, slēptās lapas un ārējās saites. Klikšķis uz rindas aizved uz šūnu; "Copy report" nokopē sarakstu."#,
        )],
    )
}

fn palette() -> Section {
    section(
        "Palete",
        vec![
            Block::Para(
                r#"Cilnē "Brand" tiek uzstādīts viss, ko pārējais panelis lieto: krāsas, fonts, valoda un valūta. Bloks "Your workbook style" rāda dzīvu paraugu ar pašreizējiem iestatījumiem."#,
            ),
            Block::Para(
                r#"Bloks "Palette" satur septiņas krāsas. Katru var izvēlēties ar krāsu izvēlni vai ierakstīt kā heksadecimālu vērtību."#,
            ),
            Block::Table {
                head: &["Krāsa", "Kur to lieto"],
                rows: &[
                    &[r#""Primary""#, "virsraksti un tumšie aizpildījumi"],
                    &[r#""Accent""#, "rezultātu rindas un izcēlumi"],
                    &[r#""Inputs""#, "ievadītie skaitļi"],
                    &[r#""Formulas""#, "aprēķinātās šūnas"],
                    &[r#""Cross-sheet links""#, "atsauces uz citu lapu"],
                    &[r#""External file links""#, "atsauces uz citu darbgrāmatu"],
                    &[r#""Partial inputs""#, "formulas, kurās ierakstīts skaitlis"],
                ],
            },
            Block::Image {
                file: "excel-brand.png",
                alt: r#"Cilne "Brand": paraugs "Your workbook style" un paletes krāsu saraksts."#,
            },
        ],
    )
}

fn logo() -> Section {
    section(
        "Krāsas no logotipa",
        vec![Block::Para(
            r#"Bloks "Upload company colors" ļauj augšupielādēt logotipu vai citu attēlu. Krāsas no tā tiek nolasītas pašā panelī, un attēls nekur netiek nosūtīts. Pirmā nospiestā krāsa kļūst par "Primary", otrā par "Accent"."#,
        )],
    )
}

fn language() -> Section {
    section(
        "Fonts, valoda un skaitļu stils",
        vec![
            Block::Para(
                r#"Blokā "Output defaults" tiek izvēlēts fonts: "Aptos", "Arial", "Calibri", "Georgia" vai "Times New Roman"."#,
            ),
            Block::Para(
                r#"Valodas izvēle nosaka mājas skaitļu stilu visam, ko panelis raksta kā tekstu."#,
            ),
            Block::Table {
                head: &["Izvēle", "Tūkstoši", "Decimāldaļa", "Valūtas zīme"],
                rows: &[
                    &[
                        r#""Latviešu""#,
                        "atdalīti ar atstarpi",
                        "punkts",
                        "aiz skaitļa",
                    ],
                    &[
                        r#""English""#,
                        "atdalīti ar komatu",
                        "punkts",
                        "pirms skaitļa",
                    ],
                    &[
                        r#""Русский""#,
                        "atdalīti ar atstarpi",
                        "punkts",
                        "aiz skaitļa",
                    ],
                ],
            },
            Block::Para(
                r#"Katrā izvēlnes rindā ir paraugs, piemēram "Latviešu: 1 094 417.5, 1 094 417 €". Decimāldaļu visās trijās valodās atdala punkts."#,
            ),
            Block::Para(
                r#"Atsevišķi no tā Excel lieto savus atdalītājus. Tas ir Excel iestatījums, ko panelis var tikai nolasīt, tāpēc, ja tas nesakrīt ar izvēlēto valodu, panelis parāda, ko Excel rāda tagad, un pasaka, kur to mainīt: Mac datorā Excel > Preferences > Edit, Windows datorā File > Options > Advanced."#,
            ),
        ],
    )
}

fn currency() -> Section {
    section(
        "Valūta",
        vec![Block::Para(
            r#"Valūtas izvēle piedāvā "€ euro", "$ dollar", "£ pound" un "No symbol". Izvēlētā zīme parādās uz valūtas pogas cilnē "Tools", valūtas formātā un valūtas ciklā. Zīmes novietojumu nosaka valoda."#,
        )],
    )
}

fn autocolor_on_edit() -> Section {
    section(
        "Autocolor on edit",
        vec![Block::Para(
            r#"Izvēles rūtiņa "Autocolor on edit" pārkrāso šūnas rediģēšanas brīdī, līdz 500 šūnām vienā reizē. Tā strādā, kamēr panelis ir atvērts, un iestatījums tiek saglabāts."#,
        )],
    )
}

fn transfer() -> Section {
    section(
        "Iestatījumu pārnešana un atiestatīšana",
        vec![
            Block::Para(
                r#"Poga "Copy palette JSON" nokopē visus iestatījumus starpliktuvē. Saite "Import JSON" tos ielasa no faila. Tā uzņēmuma paleti var nodot kolēģim vienā solī."#,
            ),
            Block::Para(
                r#"Palete tiek saglabāta gan datorā, gan pašā darbgrāmatā, tāpēc modelis saglabā savu noformējumu arī tad, kad to atver cits cilvēks citā datorā. Ja abi atšķiras, noteicošā ir darbgrāmatā saglabātā palete."#,
            ),
            Block::Para(
                r#"Poga ar apļveida bultiņu blakus virsrakstam "Palette" atgriež pls,fix noklusējuma krāsas."#,
            ),
        ],
    )
}
