from fpdf import FPDF


class ManualePDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(44, 62, 80)
        self.cell(0, 10, "CONDOMINIO ORCHIDEA - MANUALE D'USO GESTIONALE", ln=True)
        self.set_font("Helvetica", "I", 9)
        self.set_text_color(127, 140, 141)
        self.cell(0, 5, "Guida rapida all'utilizzo dell'applicazione", ln=True)
        self.ln(4)
        self.set_draw_color(44, 62, 80)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(127, 140, 141)
        self.cell(0, 10, f"Pagina {self.page_no()}", align="C")


def genera_manuale_pdf():
    pdf = ManualePDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    def aggiungi_titolo_scheda(titolo):
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_fill_color(44, 62, 80)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 8, f"  {titolo}", fill=True, ln=True)
        pdf.set_text_color(50, 50, 50)
        pdf.ln(3)

    def aggiungi_punto(testo_bold, testo_normale):
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(44, 62, 80)
        # Sostituito '•' con '-' per evitare l'errore di codifica Unicode
        pdf.write(5, f"- {testo_bold}: ")
        pdf.set_font("Helvetica", "", 9.5)
        pdf.set_text_color(60, 60, 60)
        pdf.write(5, f"{testo_normale}\n")
        pdf.ln(2)

    # SCHEDA 1
    aggiungi_titolo_scheda("1. SCHEDA: Carica Spesa / Fattura")
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(
        0,
        5,
        "Usa questa scheda ogni volta che arriva una nuova bolletta o scontrino da registrare.",
        ln=True,
    )
    pdf.ln(2)
    aggiungi_punto(
        "Tipo di Spesa",
        "Scegli tra Teleriscaldamento (ripartizione a consumi %) oppure Luce / Acqua / Spese Varie.",
    )
    aggiungi_punto(
        "Selezione Esenti",
        "Per le spese generali, puoi spuntare direttamente eventuali condomini ESENTI (es. alloggi vuoti). La spesa verra divisa solo tra i rimanenti.",
    )
    aggiungi_punto(
        "Dettagli Fattura",
        "Inserisci N° fattura, data, descrizione, importo base e periodo di consumo (Dal/Al).",
    )
    aggiungi_punto(
        "Modalita di Pagamento",
        "Selezionando Bonifico o PagoPA viene aggiunta automaticamente la commissione di EUR 1,00.",
    )
    aggiungi_punto(
        "Riferimento Pagamento",
        "Campo opzionale per inserire il codice CRO/TRN del bonifico inviato al fornitore.",
    )
    pdf.ln(4)

    # SCHEDA 2
    aggiungi_titolo_scheda("2. SCHEDA: Archivio & Modifica Spese")
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(
        0, 5, "Panoramica di tutte le spese registrate e modulo di correzione.", ln=True
    )
    pdf.ln(2)
    aggiungi_punto(
        "Tabella Generale",
        "Mostra l'importo fattura puro, la commissione, il totale e il numero di divisore applicato.",
    )
    aggiungi_punto(
        "Report Excel Teleriscaldamento",
        "Clicca sul pulsante verde per scaricare il foglio Excel con lo storico dei consumi e delle quote.",
    )
    aggiungi_punto(
        "Modifica o Elimina",
        "Seleziona la spesa dal menu a tendina, modifica i campi desiderati, spunta la casella di conferma e salva.",
    )
    pdf.ln(4)

    # SCHEDA 3
    aggiungi_titolo_scheda("3. SCHEDA: Genera Report Condomino")
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(
        0,
        5,
        "Creazione delle ricevute/avvisi di pagamento in PDF da consegnare alle famiglie.",
        ln=True,
    )
    pdf.ln(2)
    aggiungi_punto(
        "Seleziona Spese",
        "Spunta nella tabella solo le spese che vuoi includere in questo prospetto.",
    )
    aggiungi_punto(
        "Destinatario",
        "Scegli il condomino dal menu a tendina e imposta un'eventuale data di scadenza.",
    )
    aggiungi_punto(
        "Anteprima a Video",
        "Clicca su 'Genera e Visualizza Anteprima PDF' per leggere il documento direttamente nell'app.",
    )
    aggiungi_punto(
        "Download PDF",
        "Premi 'Scarica e Salva PDF' per salvartelo nella cartella desiderata del PC.",
    )
    pdf.ln(4)

    # SCHEDA 4
    aggiungi_titolo_scheda("4. SCHEDA: Incassi Condomini")
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(
        0, 5, "Registro contabile degli incassi mensili e controllo dei saldi.", ln=True
    )
    pdf.ln(2)
    aggiungi_punto(
        "Selezione Mese/Anno",
        "Il programma calcola automaticamente quanto deve versare ogni condomino nel mese.",
    )
    aggiungi_punto(
        "Gestione Esenzioni",
        "Se un alloggio era vuoto o esente, la spunta 'Esente?' azzera la sua quota a EUR 0,00.",
    )
    aggiungi_punto(
        "Stato Pagato e Note",
        "Segna con una spunta i condomini che hanno saldato e annota la data o il riferimento del bonifico.",
    )
    aggiungi_punto(
        "Riepilogo Totali",
        "Visualizza in tempo reale il Totale Spese, quanto hai incassato e quanto manca ancora.",
    )
    pdf.ln(4)

    # SCHEDA 5
    aggiungi_titolo_scheda("5. SCHEDA: Nomi Condomini")
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(0, 5, "Gestione dell'anagrafica degli appartamenti.", ln=True)
    pdf.ln(2)
    aggiungi_punto(
        "Modifica e Aggiunta",
        "Permette di rinominare i condomini esistenti o aggiungerne di nuovi con il pulsante '+ Aggiungi'.",
    )

    pdf.output("Manuale_Uso_Condominio_Orchidea.pdf")
    print("✅ PDF del manuale generato con successo!")


if __name__ == "__main__":
    genera_manuale_pdf()
