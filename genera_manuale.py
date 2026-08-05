from fpdf import FPDF


class PDFManuale(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 15)
        self.set_text_color(44, 62, 80)
        self.cell(0, 10, "CONDOMINIO ORCHIDEA - MANUALE D'USO", ln=True, align="C")
        self.set_font("Helvetica", "I", 10)
        self.set_text_color(127, 140, 141)
        self.cell(
            0,
            5,
            "Guida rapida all'utilizzo del Gestionale (Versione Aggiornata)",
            ln=True,
            align="C",
        )
        self.ln(4)
        self.set_draw_color(200, 200, 200)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Pagina {self.page_no()}", align="C")


pdf = PDFManuale()
pdf.add_page()
pdf.set_auto_page_break(auto=True, margin=15)

sezioni = [
    (
        "1. SCHEDA: Carica Spesa / Fattura",
        [
            "Usa questa scheda ogni volta che arriva una nuova bolletta o scontrino da registrare.",
            "- Tipo di Spesa: Scegli tra Teleriscaldamento (ripartizione a consumi %) oppure Luce / Acqua / Spese Varie.",
            "- Selezione Esenti: Per le spese generali, puoi spuntare direttamente eventuali condomini ESENTI (es. alloggi vuoti). La spesa verra divisa solo tra i rimanenti.",
            "- Dettagli Fattura: Inserisci N° fattura, data, descrizione, importo base e periodo di consumo (Dal / Al).",
            "- Modalita di Pagamento: Selezionando Bonifico o PagoPA viene aggiunta automaticamente la commissione di EUR 1,00.",
            "- Riferimento Pagamento: Campo opzionale per inserire il codice CRO/TRN/PagoPA del bonifico inviato al fornitore.",
            "- Protezione Salvataggio (Anti-Doppio Click): Una volta cliccato su 'Salva Spesa in Archivio', il salvataggio avviene all'istante e il pulsante si disattiva per evitare di archiviare due volte la stessa spesa per sbaglio. Per sbloccare la maschera e inserire una nuova spesa, basta cliccare sul pulsante 'Inserisci un'altra spesa'.",
        ],
    ),
    (
        "2. SCHEDA: Archivio & Modifica Spese",
        [
            "Panoramica di tutte le spese registrate e modulo di correzione.",
            "- Tabella Generale: Mostra l'importo fattura puro, la commissione, il totale e il numero di divisore applicato.",
            "- Report Excel Teleriscaldamento: Clicca sul pulsante per scaricare il foglio Excel con lo storico dei consumi e delle quote.",
            "- Modifica o Elimina: Seleziona la spesa dal menu a tendina, modifica i campi desiderati, spunta la casella di conferma e salva (oppure eliminala).",
        ],
    ),
    (
        "3. SCHEDA: Genera Report Condomino",
        [
            "Creazione delle ricevute/avvisi di pagamento in PDF da consegnare alle famiglie.",
            "- Filtro Automatico Spese (Novita): All'apertura della scheda, le spese contrassegnate come 'GIA REPORTATA' vengono deselezionate automaticamente (senza spunta), cosi da evitare di inserire per errore vecchie spese gia consegnate. Saranno preselezionate solo le spese 'NUOVE'. Se desideri includere nuovamente una spesa passata, ti bastera spuntarla manualmente.",
            "- Destinatario: Scegli il condomino dal menu a tendina e imposta un'eventuale data di scadenza.",
            "- Anteprima a Video: Clicca su 'Genera e Visualizza Anteprima PDF' per leggere il documento direttamente nell'app.",
            "- Download PDF: Premi 'Scarica e Salva PDF' per salvarlo nella cartella desiderata del PC.",
            "- Chiusura Spese: Clicca su 'Segna spese come GIA REPORTATE' per aggiornare lo stato delle spese incluse nel prospetto.",
        ],
    ),
    (
        "4. SCHEDA: Incassi Condomini",
        [
            "Registro contabile degli incassi mensili e controllo dei saldi.",
            "- Selezione Mese/Anno: Il programma calcola automaticamente quanto deve versare ogni condomino nel mese.",
            "- Gestione Esenzioni: Se un alloggio era vuoto o esente, la spunta 'Esente?' azzera la sua quota a EUR 0,00.",
            "- Stato Pagato e Note: Segna con una spunta i condomini che hanno saldato e annota la data o il riferimento del bonifico.",
            "- Riepilogo Totali: Visualizza in tempo reale il Totale Spese, quanto hai incassato e quanto manca ancora.",
        ],
    ),
    (
        "5. SCHEDA: Nomi Condomini",
        [
            "Gestione dell'anagrafica degli appartamenti.",
            "- Modifica e Aggiunta: Permette di rinominare i condomini esistenti o aggiungerne di nuovi con il pulsante '+ Aggiungi'.",
        ],
    ),
]

for titolo, righe in sezioni:
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(44, 62, 80)
    pdf.cell(0, 8, titolo.encode("latin-1", "replace").decode("latin-1"), ln=True)
    pdf.set_font("Helvetica", "", 9.5)
    pdf.set_text_color(50, 50, 50)
    for riga in righe:
        pdf.multi_cell(0, 5, riga.encode("latin-1", "replace").decode("latin-1"))
        pdf.ln(1)
    pdf.ln(2)

pdf.output("Manuale_Uso_Condominio_Orchidea.pdf")
print("✅ File 'Manuale_Uso_Condominio_Orchidea.pdf' generato con successo!")
