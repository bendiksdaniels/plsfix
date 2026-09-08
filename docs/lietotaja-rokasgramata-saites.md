# Excel un PowerPoint saistītie objekti (v2.1)

pls,fix ļauj eksportēt Excel šūnu apgabalu vai diagrammu uz PowerPoint kā saistītu attēlu, rediģējamu tabulu vai diagrammu (figūru grupu). Objektu pēc tam var atjaunināt vietā, nemainot tā pozīciju un izmēru, arī tad, kad avota dati Excel darbgrāmatā ir mainījušies. Šī rokasgrāmata apraksta eksportu, ievietošanu, atjaunināšanu, saites atslēgas iestatīšanu, drošības prasības un biežākās kļūdas.

## Pirms sākt

- Nepieciešams Microsoft 365: PowerPoint 2504 vai jaunāka versija Windows datoram, PowerPoint 16.96 vai jaunāka versija Mac datoram.
- Saites atslēga jāievada vienreiz katrā datorā: izveidojiet to Excel pusē un ielīmējiet PowerPoint pusē (skatiet sadaļu "Saites atslēga").

## Eksports no Excel

Excel cilnē "Links" atlasiet šūnu apgabalu vai diagrammu, ko vēlaties nosūtīt uz PowerPoint.

- Šūnu apgabalam nospiediet "Export selection".
- Diagrammai nospiediet "Export active chart".
- Ar "Export as table" tas pats šūnu apgabals tiek nosūtīts kā rediģējama PowerPoint tabula (līdz 60 rindām un 20 kolonnām).
- *Export as text* nosūta vienas šūnas attēloto tekstu (skaitli vai frāzi) kā teksta lauku, ko PowerPoint atjauno tajā pašā vietā; lauka izmērs un fonts paliek tādi, kādus tos atstājāt (līdz 500 rakstzīmēm).

Excel izveido avotam slēptu nosaukumu, kas seko tam arī tad, kad virs apgabala tiek ievietotas rindas vai lapa tiek pārdēvēta. Attēls tiek nosūtīts uz releju un PowerPoint pusē parādās kā gaidošs vienums.

Izvēles rūtiņa "Highlight linked cells" ar vieglu toni iekrāso visus saistītos šūnu apgabalus un, izslēdzot to, atjauno sākotnējo noformējumu, savukārt diagrammas netiek iekrāsotas.

## Ievietošana PowerPoint

PowerPoint lentē atveriet cilni "pls,fix" un nospiediet "Links": atvērsies saišu panelis. Pārslēdzieties uz cilni "Inbox", kur redzami no Excel nosūtītie, vēl neievietotie attēli.

Atlasiet slaidu, kurā attēls jāievieto, un pie vajadzīgā vienuma nospiediet "Insert". Attēls tiek ievietots atlasītajā slaidā un pielāgots tā izmēram.

## Atjaunināšana

Kad avota dati Excel mainās, saites jāatjaunina abās pusēs.

- Excel pusē cilnē "Links" nospiediet "Push all": visas saites tiek pārrēķinātas no to avotiem un nosūtītas uz releju.
- PowerPoint pusē cilnē "Links" nospiediet "Update all": visi izsekotie attēli tiek pārzīmēti vietā. Var atjaunināt arī atsevišķas saites ("Update selected") vai visas aktīvā slaida saites ("Update this slide").

Ja Excel cilnē "Links" ir ieslēgta izvēles rūtiņa "Auto-push on edit", mainītās saites tiek nosūtītas uz releju automātiski trīs sekundes pēc pēdējās rediģēšanas, kamēr panelis ir atvērts.

Pozīcija un platums saglabājas nemainīgi. Augstums mainās tikai tad, kad attēla proporcijas Excel pusē ir mainījušās. Ja attēls ir ievietots grupā, atjaunināšana to tomēr atrod un atjauno.

Lai attēlu saistītu ar citu eksportu, piemēram, ar to pašu tabulu no jaunākas darbgrāmatas, atzīmējiet vienu rindu, nospiediet "Change source" un izvēlieties kādu no cilnē "Inbox" gaidošajiem eksportiem; attēla slaids, pozīcija un izmērs paliek nemainīgi.

## Saites atslēga

Saite starp Excel un PowerPoint darbojas tikai tad, kad abām pusēm ir viena un tā pati atslēga.

1. Excel cilnē "Links" sadaļā "Link key" nospiediet "Generate", lai izveidotu jaunu atslēgu.
2. Nospiediet "Copy", lai to nokopētu.
3. PowerPoint cilnē "Settings" ielīmējiet atslēgu un nospiediet "Save key".

Šī darbība jāveic tikai vienreiz katrā datorā. Jauna atslēga neietekmē jau ievietotās saites, bet PowerPoint ar veco atslēgu vairs neredz jaunus eksportus, kamēr tajā nav ielīmēta jaunā.

## Drošība

Relejs glabā eksportētos attēlus tikai šifrētā veidā un ne ilgāk kā 30 dienas. Ikviens, kam ir pieejama prezentācija, šajā laikā var lejupielādēt saites jaunāko attēlu, pat ja saite prezentācijā vairs nav redzama. Tāpēc pirms prezentācijas nosūtīšanas ārpus bankas katrai saitei PowerPoint cilnē "Links" jānospiež "Break link".

## Biežākās kļūdas

| Ziņojums | Cēlonis | Ko darīt |
|---|---|---|
| "Source missing" | Excel pusē avota nosaukums vai diagramma ir dzēsta. | Excel cilnē "Links" pārbaudiet rindu ar "Go to source" un, ja nepieciešams, eksportējiet saiti no jauna. |
| "Wrong link key" | Attēlā saglabātais saites marķieris neatbilst relejā reģistrētajam (piemēram, attēla kopija ar bojātu marķieri). | Excel cilnē "Links" eksportējiet saiti no jauna un ievietojiet attēlu vēlreiz no "Inbox". |
| "Select a slide first." | PowerPoint slaidu rādītājā nav atlasīts neviens slaids. | Atlasiet slaidu, kurā jāievieto vai jāatjaunina attēls, un mēģiniet vēlreiz. |
| "select a single range" | Excel atlasē ir vairāki nesaistīti apgabali, piemēram, atlasīti ar Ctrl taustiņu. | Atlasiet vienu nepārtrauktu šūnu apgabalu un nospiediet "Export selection" vēlreiz. |

## Ierobežojumi

- Ctrl+Z neatceļ pievienojumprogrammas veiktās izmaiņas, tāpēc pēc "Update all" atzīmētās rindas var atgriezt uz iepriekšējo versiju ar pogu "Revert last update", taču relejs glabā tikai vienu iepriekšējo versiju, tāpēc tālāk atpakaļ atgriezties nav iespējams.
- Tabula nonāk kā rediģējama PowerPoint tabula (līdz 60 rindām un 20 kolonnām), diagramma kā figūru grupa. Diagramma, ko figūras neatveido (vairāk nekā 6 sērijas, vairāk nekā 40 punkti vai 12 sektori, neatbalstīts diagrammas tips), nonāk kā attēls, un abi paneļi pasaka iemeslu.
- Darbgrāmatas versiju maiņa notiek tikai caur cilni "Inbox": vispirms jaunā darbgrāmata jāeksportē no Excel, un tikai pēc tam saiti var pārvirzīt uz to.
