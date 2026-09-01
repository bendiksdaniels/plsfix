# pls,fix IT vienlapa

pls,fix ir Excel un PowerPoint pievienojumprogramma finanšu modelēšanai. Šī lapa apkopo to, kas IT nodaļai jāzina pirms centralizētas izplatīšanas Microsoft 365 vidē.

## Kas tas ir

- pls,fix ir uzdevumu paneļa (task pane) tipa pievienojumprogramma diviem Office produktiem: Excel un PowerPoint.
- Tā ir uzbūvēta uz Office.js, Microsoft oficiālās pievienojumprogrammu platformas; kods darbojas Office iekšējā pārlūkā, nevis kā atsevišķa lietotne.
- Pievienojumprogramma nesatur makro vai VBA kodu un nemaina darbgrāmatas vai prezentācijas drošības iestatījumus.
- Lietotāja datorā nekas netiek instalēts, izņemot manifesta failu, kas pasaka Office, no kurienes paneli ielādēt.

## Izvietošana

- Administrators Microsoft 365 administrācijas centrā atver "Settings" > "Integrated apps" > "Upload custom apps", augšupielādē manifesta failu manifest.prod.xml un piešķir to lietotāju grupai.
- Programmas koda (JS/HTML/CSS) atjauninājumi nesasniedz administratoru: tie notiek serverī, un lietotājs tos redz nākamajā palaišanas reizē.
- Manifesta izmaiņas, piemēram, jauna poga lentē, prasa atkārtotu augšupielādi administrācijas centrā.
- Funkcijas =PLSFIX.ROUND un =PLSFIX.ROUNDSUM Office kešatmiņā glabājas atsevišķi no pārējā koda; to atjauninājums pie lietotājiem var nonākt līdz 24 stundām pēc izvietošanas.

## Kur tas darbojas

- Panelis ielādējas no https://dbautomatizacijas.com/modelis/.
- Adrese darbojas uz Hetzner VPS aiz Cloudflare, tajā pašā serverī, kas apkalpo visu DB Automatizācijas rīku komplektu.
- Šim ceļam ir Cloudflare Access apiešana (bypass): Office pievienojumprogrammu iekšējais logs nevar izpildīt Access pieteikšanos, tāpēc piekļuvi nosaka tikai Microsoft 365 centralizētās izplatīšanas piešķīrums.

## Dati un drošība

- Darbgrāmatas un prezentācijas dati serverim nekad nenonāk; vienīgais, kas ceļo, ir PowerPoint saistīto objektu attēls.
- Šis attēls tiek šifrēts pašā panelī (AES-GCM) pirms nosūtīšanas; atslēgu atvasina no saites atslēgas, kas atrodas tikai prezentācijas slaidu birkās un darbgrāmatas reģistrā.
- Serveris glabā tikai šifrētu saturu un autorizācijas atslēgas SHA-256 nospiedumu, nevis pašu atslēgu, failu nosaukumus vai saturu; nav lietotāju kontu un nav telemetrijas.
- Saites derīgas 30 dienas, "Inbox" vienumi 7 dienas; pēc tam tie no servera tiek dzēsti.
- Serveris ierobežo pieprasījumu skaitu katram klientam (300 rakstīšanas un 1 200 lasīšanas pieprasījumi minūtē), kopējo glabāto apjomu visiem kopā (1 GiB) un "Inbox" vienumu skaitu katrai saites atslēgai (500).

## Prasības

- Microsoft 365 Apps, Current Channel, gan Excel, gan PowerPoint.
- Excel: ExcelApi 1.9 vai jaunāka (iekļauta ikvienā pašreizējā Microsoft 365 versijā).
- PowerPoint: 2021 vai Microsoft 365 versija tabulu, diagrammu un "Tools" cilnes rīku izmantošanai.
- Tīkla pieeja dbautomatizacijas.com pa HTTPS; cita veida piekļuve (VPN, atsevišķa pieteikšanās) nav nepieciešama.

## Atbalsts

- Kontaktpersona: Daniels Bendiks.
- Uzstādītā versija redzama paneļa apakšējā labajā stūrī.
- Servera versiju var pārbaudīt jebkurā pārlūkā: https://dbautomatizacijas.com/modelis/version.
