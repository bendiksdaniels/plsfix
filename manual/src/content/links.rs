// Chapter 6: linked objects between Excel and PowerPoint. Reuses the wording of
// docs/lietotaja-rokasgramata-saites.md and owns both sides of the flow, the
// Excel "Links" tab and the four PowerPoint tabs.

use super::{Block, Chapter, Section, section};

pub fn links() -> Chapter {
    Chapter {
        title: "Saites uz PowerPoint",
        sections: vec![
            how_it_works(),
            link_key(),
            export(),
            list(),
            toggles(),
            insert(),
            update(),
            filters(),
            manage(),
            object_tools(),
            smart_painter(),
            security(),
        ],
    }
}

fn how_it_works() -> Section {
    section(
        "Kā saite darbojas",
        vec![
            Block::Para(
                r#"pls,fix ļauj eksportēt Excel šūnu apgabalu vai diagrammu uz PowerPoint kā saistītu attēlu, tabulu vai diagrammu. Objektu pēc tam var atjaunināt vietā, nemainot tā pozīciju un izmēru, arī tad, kad avota dati Excel darbgrāmatā ir mainījušies."#,
            ),
            Block::Para(
                r#"Excel puse ieraksta avotam slēptu nosaukumu, tāpēc saite seko līdzi arī tad, kad virs apgabala tiek ievietotas rindas vai lapa tiek pārdēvēta. Attēls tiek nosūtīts uz releju, no kurienes to paņem PowerPoint puse."#,
            ),
            Block::Para(
                r#"Saitē var būt attēls, tabula vai diagramma. Diagramma slaidā nonāk kā rediģējamu figūru grupa ar zīmola krāsām un tikai vērtību etiķetēm: stabiņu, joslu, tilta, sektoru un līniju diagrammas. Līniju diagrammā katrs punkts ir sava figūra, savienota ar līniju līdz nākamajam punktam. Šai figūru grupai nepieciešams PowerPoint 2504 vai jaunāks Windows datorā, 16.96 vai jaunāks Mac datorā; vecākā versijā, kā arī pāri 40 punktiem, 3 sērijām vai 12 sektoriem, diagramma tiek ievietota kā attēls, un panelis to pasaka."#,
            ),
        ],
    )
}

fn link_key() -> Section {
    section(
        "Saites atslēga",
        vec![
            Block::Para(
                r#"Saite starp Excel un PowerPoint darbojas tikai tad, kad abām pusēm ir viena un tā pati atslēga. Tā jāievada vienreiz katrā datorā."#,
            ),
            Block::Steps(&[
                r#"Excel cilnē "Links" sadaļā "Link key" nospiediet "Generate", lai izveidotu jaunu atslēgu."#,
                r#"Nospiediet "Copy", lai to nokopētu. Panelī ir redzami tikai atslēgas gali, bet "Copy" nodod visu atslēgu."#,
                r#"PowerPoint cilnē "Settings" ielīmējiet atslēgu un nospiediet "Save key"."#,
            ]),
            Block::Para(
                r#"Kad atslēga ir saglabāta, PowerPoint pusē redzams "Paired". Poga "Forget key" to izdzēš no šī datora, bet poga "Reveal" Excel pusē parāda atslēgu pilnībā."#,
            ),
            Block::Para(
                r#"Jauna atslēga neietekmē jau ievietotās saites, bet PowerPoint ar veco atslēgu vairs neredz jaunus eksportus, kamēr tajā nav ielīmēta jaunā. Atslēga paliek datorā, un relejs to nekad neredz."#,
            ),
            Block::Image {
                file: "ppt-settings.png",
                alt: r#"PowerPoint cilne "Settings" ar saites atslēgas lauku un pogām "Save key" un "Forget key"."#,
            },
        ],
    )
}

fn export() -> Section {
    section(
        "Eksports no Excel",
        vec![
            Block::Para(r#"Excel cilnē "Links" atlasiet to, ko vēlaties nosūtīt uz PowerPoint."#),
            Block::Bullets(&[
                r#"Šūnu apgabalam nospiediet "Export selection". Atlasei jābūt vienam nepārtrauktam apgabalam."#,
                r#"Diagrammai nospiediet "Export active chart"."#,
            ]),
            Block::Para(
                r#"Ja diagramma nav atlasīta, panelis paņem lapas vienīgo diagrammu. Ja lapā ir vairākas, zem pogas parādās diagrammu saraksts, no kura izvēlēties vajadzīgo."#,
            ),
            Block::Para(
                r#"Ar "Export as text" vienas šūnas attēlotais teksts nonāk slaidā kā atsevišķs teksta lauks. Atjaunināšana maina tikai tekstu; lauka vieta, izmērs un fonts paliek tādi, kādus tos atstājāt (līdz 500 rakstzīmēm)."#,
            ),
            Block::Image {
                file: "excel-links.png",
                alt: r#"Excel cilne "Links": pogas "Export selection" un "Export active chart", diagrammu saraksts un saišu saraksts."#,
            },
        ],
    )
}

fn list() -> Section {
    section(
        "Saišu saraksts Excel pusē",
        vec![
            Block::Para(
                r#"Bloks "Linked objects" rāda visas šīs darbgrāmatas saites: avotu un laiku, kad saite pēdējo reizi nosūtīta."#,
            ),
            Block::Table {
                head: &["Poga", "Ko tā dara"],
                rows: &[
                    &[r#""Push selected""#, "pārrēķina un nosūta atzīmētās saites"],
                    &[r#""Push all""#, "pārrēķina un nosūta visas saites"],
                    &[
                        r#""Go to source""#,
                        "pārlec uz atzīmētās saites avotu darbgrāmatā",
                    ],
                    &[
                        r#""Remove link""#,
                        "izņem saiti no reģistra; jau ievietotais attēls prezentācijā paliek",
                    ],
                ],
            },
        ],
    )
}

fn toggles() -> Section {
    section(
        "Auto-push on edit un Highlight linked cells",
        vec![
            Block::Para(
                r#"Izvēles rūtiņa "Auto-push on edit" nosūta mainītās saites uz releju automātiski trīs sekundes pēc pēdējās rediģēšanas. Tā strādā, kamēr panelis ir atvērts, un tiek iegaumēta katrai darbgrāmatai atsevišķi."#,
            ),
            Block::Para(
                r#"Izvēles rūtiņa "Highlight linked cells" ar vieglu toni iekrāso visus saistītos šūnu apgabalus, tāpēc uzreiz redzams, kas baro prezentāciju. Izslēdzot to, sākotnējais noformējums tiek atjaunots. Diagrammas netiek iekrāsotas."#,
            ),
        ],
    )
}

fn insert() -> Section {
    section(
        "Ievietošana PowerPoint",
        vec![
            Block::Para(
                r#"PowerPoint lentē atveriet cilni "pls,fix" un nospiediet "Links": atvērsies saišu panelis. Pārslēdzieties uz cilni "Inbox", kur redzami no Excel nosūtītie, vēl neievietotie attēli."#,
            ),
            Block::Steps(&[
                "Atlasiet slaidu, kurā attēls jāievieto.",
                r#"Pie vajadzīgā vienuma nospiediet "Insert"."#,
                "Objekts (attēls, tabula vai diagramma) tiek ievietots atlasītajā slaidā brīvā vietā un pielāgots tā izmēram. Pēc tam to var pārvietot un mainīt tā izmēru.",
            ]),
            Block::Para(
                r#"Poga "Paste latest linked" paveic to pašu vienā solī: tā ievieto jaunāko no Excel nosūtīto eksportu aktīvajā slaidā, neizvēloties konkrētu vienumu sarakstā."#,
            ),
            Block::Para(
                r#"Vienumi cilnē "Inbox" ir derīgi 7 dienas pēc nosūtīšanas. Ja saraksts ir tukšs un panelī redzams "Not paired", vispirms ielīmējiet saites atslēgu cilnē "Settings"."#,
            ),
            Block::Image {
                file: "ppt-inbox.png",
                alt: r#"PowerPoint cilne "Inbox" ar tukšu sarakstu un paskaidrojumu par vienumu derīguma termiņu."#,
            },
        ],
    )
}

fn update() -> Section {
    section(
        "Atjaunināšana",
        vec![
            Block::Para(
                r#"Kad avota dati Excel mainās, saites jāatjaunina abās pusēs. Excel pusē cilnē "Links" nospiediet "Push all", pēc tam PowerPoint pusē cilnē "Links" nospiediet vienu no atjaunināšanas pogām."#,
            ),
            Block::Table {
                head: &["Poga", "Ko tā atjaunina"],
                rows: &[
                    &[r#""Update selected""#, "atzīmētās saites"],
                    &[r#""Update this slide""#, "visas aktīvā slaida saites"],
                    &[r#""Update all""#, "visas prezentācijas saites"],
                    &[
                        r#""Revert last update""#,
                        "atgriež atzīmētās saites uz iepriekšējo attēlu",
                    ],
                ],
            },
            Block::Para(
                r#"Pozīcija un platums saglabājas nemainīgi. Augstums mainās tikai tad, kad attēla proporcijas Excel pusē ir mainījušās. Ja attēls ir ievietots grupā, atjaunināšana to tomēr atrod un atjauno."#,
            ),
            Block::Para(
                r#"Relejs glabā vienu iepriekšējo versiju, tāpēc "Revert last update" atgriež vienu soli atpakaļ, bet ne tālāk."#,
            ),
            Block::Image {
                file: "ppt-links.png",
                alt: r#"PowerPoint cilne "Links": atjaunināšanas pogas un saišu tabula ar slaidu, avotu un stāvokli."#,
            },
        ],
    )
}

fn filters() -> Section {
    section(
        "Meklēšana un filtrēšana",
        vec![
            Block::Para(
                r#"Virs saišu tabulas cilnē "Links" ir meklēšanas lauks un trīs filtri. Tie sašaurina redzamo sarakstu, bet atzīmētās rindas paliek atzīmētas arī tad, kad filtrs tās paslēpj."#,
            ),
            Block::Table {
                head: &["Vadīkla", "Ko tā sašaurina"],
                rows: &[
                    &[
                        "Meklēšanas lauks",
                        "meklē pēc slaida numura, objekta un tā avota darbgrāmatas nosaukuma",
                    ],
                    &[
                        r#""Source workbook""#,
                        "rāda tikai izvēlētās darbgrāmatas saites",
                    ],
                    &[r#""Slide""#, "rāda tikai izvēlētā slaida saites"],
                    &[r#""Link status""#, "rāda tikai izvēlētā stāvokļa saites"],
                ],
            },
            Block::Para("Stāvokļa filtra četras vērtības:"),
            Block::Bullets(&[
                r#""Needs update": avota dati Excel pusē mainījušies kopš pēdējās atjaunināšanas."#,
                r#""Missing source": relejā šai saitei vairs nav neviena attēla."#,
                r#""Wrong link key": attēls saglabāts ar citu saites atslēgu, nevis šī datora aktuālo."#,
                r#""Up to date": attēls atbilst jaunākajam Excel eksportam."#,
            ]),
        ],
    )
}

fn manage() -> Section {
    section(
        "Go to slide, Change source un Break link",
        vec![
            Block::Para(r#"Poga "Go to slide" pārlec uz pirmās atzīmētās saites slaidu."#),
            Block::Para(
                r#"Lai attēlu saistītu ar citu eksportu, piemēram, ar to pašu tabulu no jaunākas darbgrāmatas, atzīmējiet vienu rindu, nospiediet "Change source" un izvēlieties kādu no cilnē "Inbox" gaidošajiem eksportiem. Izvēli apstiprina poga "Confirm", atceļ poga "Cancel". Attēla slaids, pozīcija un izmērs paliek nemainīgi."#,
            ),
            Block::Para(
                r#"Darbgrāmatas versiju maiņa notiek tikai caur cilni "Inbox": vispirms jaunā darbgrāmata jāeksportē no Excel, un tikai pēc tam saiti var pārvirzīt uz to."#,
            ),
            Block::Para(
                r#"Poga "Break link" noņem attēlam saites marķieri. Attēls slaidā paliek, bet vairs netiek atjaunināts."#,
            ),
        ],
    )
}

fn object_tools() -> Section {
    section(
        "Object tools",
        vec![
            Block::Para(
                r#"PowerPoint cilnē "Tools" ir sadaļa "Object tools" darbam ar slaidā atlasītajām figūrām. Tā prasa PowerPoint 2021 vai Microsoft 365; vecākā versijā panelis to pasaka, tiklīdz nospiež kādu no šīm pogām. Tie paši rīki ir arī PowerPoint lentes cilnē "pls,fix": grupā "Objects" (Object tools, Match size, Select similar, Swap, Capture style, Apply style) un grupā "Arrange" (sešas līdzināšanas un abas izkārtošanas), katrs ar savu ikonu."#,
            ),
            Block::Para(
                r#""Align" izvēlnē izvēlieties virzienu (pa kreisi, centrā, pa labi, augšā, vidū vai apakšā) un nospiediet "Apply": atlasītās figūras līdzinās pēc visu atlasīto figūru kopējām malām. Nepieciešamas vismaz divas figūras."#,
            ),
            Block::Para(
                r#""Distribute" izvēlnē izvēlieties virzienu (šķērsām vai lejup) un nospiediet "Apply": atstarpes starp trim vai vairāk figūrām kļūst vienādas; pirmās un pēdējās figūras vieta nemainās. Nepieciešamas vismaz trīs figūras."#,
            ),
            Block::Table {
                head: &["Poga", "Ko tā dara"],
                rows: &[
                    &[
                        r#""Match size""#,
                        "uzliek pirmās atlasītās figūras platumu un augstumu visām pārējām atlasītajām figūrām",
                    ],
                    &[
                        r#""Select similar""#,
                        "no vienas atlasītas figūras atlasa visas pārējās tā paša tipa un gandrīz tāda paša izmēra figūras tajā pašā slaidā",
                    ],
                    &[r#""Swap""#, "samaina vietām tieši divas atlasītas figūras"],
                ],
            },
        ],
    )
}

fn smart_painter() -> Section {
    section(
        "Smart Painter",
        vec![
            Block::Para(
                r#"Zem "Object tools" ir "Smart Painter": tas nokopē vienas figūras aizpildījumu un kontūru uz citām figūrām, neatkarīgi no to izmēra vai satura."#,
            ),
            Block::Steps(&[
                r#"Atlasiet vienu figūru ar vienkrāsainu aizpildījumu vai bez aizpildījuma un nospiediet "Capture"."#,
                "Atlasiet vienu vai vairākas mērķa figūras.",
                r#"Nospiediet "Apply"."#,
            ]),
            Block::Para(
                r#"Attēla vai gradienta aizpildījumu "Capture" nepieņem: der tikai vienkrāsains aizpildījums vai figūra bez aizpildījuma."#,
            ),
        ],
    )
}

fn security() -> Section {
    section(
        "Drošība",
        vec![
            Block::Para(
                r#"Relejs glabā eksportētos attēlus tikai šifrētā veidā un ne ilgāk kā 30 dienas. Ikviens, kam ir pieejama prezentācija, šajā laikā var lejupielādēt saites jaunāko attēlu, pat ja saite prezentācijā vairs nav redzama."#,
            ),
            Block::Para(
                r#"Tāpēc pirms prezentācijas nosūtīšanas ārpus organizācijas katrai saitei PowerPoint cilnē "Links" jānospiež "Break link"."#,
            ),
        ],
    )
}
