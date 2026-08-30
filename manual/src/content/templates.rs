// Chapter 5: the "Templates" section at the top of the Tools tab, right under
// the Selection inspector, one section per ready block. Owns what each template writes and how big its block is; the
// button labels stay in English exactly as the pane shows them.

use super::{Block, Chapter, Section, section};

pub fn templates() -> Chapter {
    Chapter {
        title: "Veidnes",
        sections: vec![
            overview(),
            debt_schedule(),
            dcf(),
            npv_irr(),
            working_capital(),
            sensitivity(),
            ebitda_bridge(),
        ],
    }
}

fn overview() -> Section {
    section(
        "Kas ir veidnes",
        vec![
            Block::Para(
                r#"Cilnes "Tools" sākumā, tūlīt zem atlases pārskata, ir sadaļa "Templates" ar sešām pogām. Katra poga ieraksta gatavu aprēķina bloku, kura augšējais kreisais stūris ir aktīvajā šūnā. Bloku var ierakstīt jebkurā lapā un jebkurā vietā, jo visas formulas norāda uz paša bloka šūnām."#,
            ),
            Block::Para(
                r#"Ja bloka vietā jau kaut kas atrodas, panelis to neaizstāj, bet paziņo, cik rindu un cik kolonnu brīvas vietas tam vajag."#,
            ),
            Block::Para(
                r#"Ierakstītie skaitļi ir zili, formulas melnas, virsraksta un galvenes rindas noformētas jūsu paletē, un skaitļu formāti nāk no paneļa iestatījumiem. Pēc ierakstīšanas bloks paliek atlasīts, un to var atcelt ar pogu "Undo last pls,fix action"."#,
            ),
        ],
    )
}

fn debt_schedule() -> Section {
    section(
        "Debt schedule (annuity)",
        vec![Block::Para(
            r#"Poga "Debt schedule (annuity)" ieraksta 20 rindu un 6 kolonnu bloku. Ievaddati ir aizdevuma summa, gada likme, gadu skaits un maksājumu skaits gadā; no tiem bloks aprēķina periodu skaitu un vienādu periodisko maksājumu ar funkciju PMT. Desmit periodu rindās ir sākuma atlikums, maksājums, procenti, pamatsummas daļa un beigu atlikums, un pēdējā rinda saskaita maksājumus, procentus un pamatsummu."#,
        )],
    )
}

fn dcf() -> Section {
    section(
        "DCF valuation",
        vec![Block::Para(
            r#"Poga "DCF valuation" ieraksta 13 rindu un 6 kolonnu bloku. Ievaddati ir WACC, terminālā izaugsme un piecu gadu brīvā naudas plūsma. Bloks aprēķina katra gada diskonta koeficientu, naudas plūsmas pašreizējo vērtību, terminālo vērtību pēc Gordona formulas, tās pašreizējo vērtību un uzņēmuma vērtību."#,
        )],
    )
}

fn npv_irr() -> Section {
    section(
        "NPV / IRR",
        vec![Block::Para(
            r#"Poga "NPV / IRR" ieraksta 17 rindu un 3 kolonnu bloku. Ievaddati ir sākotnējais ieguldījums, astoņas periodu naudas plūsmas un diskonta likme. Bloks veido uzkrātās naudas plūsmas kolonnu un aprēķina neto pašreizējo vērtību, iekšējo ienesīguma likmi un atmaksāšanās periodu. Atmaksāšanās periods arī ir formula, nevis ierakstīts skaitlis, un tas rāda "n/a", ja ieguldījums neatmaksājas."#,
        )],
    )
}

fn working_capital() -> Section {
    section(
        "Working capital days",
        vec![Block::Para(
            r#"Poga "Working capital days" ieraksta 12 rindu un 2 kolonnu bloku. Ievaddati ir apgrozījums, pārdotās produkcijas izmaksas, debitori, krājumi un kreditori. Bloks aprēķina debitoru, krājumu un kreditoru aprites dienas un naudas konversijas ciklu."#,
        )],
    )
}

fn sensitivity() -> Section {
    section(
        "Two-way sensitivity",
        vec![Block::Para(
            r#"Poga "Two-way sensitivity" ieraksta 11 rindu un 6 kolonnu bloku. Ievaddati ir bāzes vērtība un divu faktoru soļa lielumi. Bloks izveido 5 x 5 formulu tabulu, kurā katra šūna reizina bāzes vērtību ar rindas un kolonnas soli. Asu vērtības arī ir formulas, tāpēc soļa maiņa pārrēķina visu tabulu; Excel iebūvētā Data Table funkcija netiek izmantota."#,
        )],
    )
}

fn ebitda_bridge() -> Section {
    section(
        "EBITDA bridge",
        vec![Block::Para(
            r#"Poga "EBITDA bridge" ieraksta 9 rindu un 2 kolonnu bloku: sākuma EBITDA, pieci nosaukti soļi, beigu EBITDA un pārbaudes rinda. Bloks ir tieši tādā formā, kādu gaida poga "Waterfall from selection": nosaukumi kreisajā kolonnā, skaitļi labajā, kopsummas pirmajā un pēdējā rindā. Diagrammai jāatlasa septiņas tilta rindas bez virsraksta rindas un bez pārbaudes rindas."#,
        )],
    )
}
