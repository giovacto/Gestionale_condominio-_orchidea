import ast
import base64
from datetime import datetime
import io
import sqlite3
from fpdf import FPDF
import pandas as pd
import streamlit as st

# Impostazione della pagina
st.set_page_config(page_title="Condominio Orchidea", layout="wide", page_icon="🌺")


# ==========================================
# FUNZIONI HELPER
# ==========================================
def mostra_pdf_in_anteprima(pdf_bytes):
    base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)


def formatta_data_ita(data_str):
    if not data_str:
        return ""
    try:
        if "-" in data_str and len(data_str) == 10:
            parti = data_str.split("-")
            return f"{parti[2]}/{parti[1]}/{parti[0]}"
    except Exception:
        pass
    return data_str


def parse_data_ita(data_str):
    try:
        return datetime.strptime(data_str, "%d/%m/%Y").date()
    except Exception:
        return datetime.now().date()


def pulisci_testo_pdf(testo):
    if not testo:
        return ""
    return str(testo).replace("€", "EUR")


def genera_excel_teleriscaldamento(df_spese_tele, list_condomini):
    righe_excel = []
    for _, row in df_spese_tele.iterrows():
        dett_perc = {}
        try:
            if row["dettagli_json"]:
                dett_perc = ast.literal_eval(row["dettagli_json"])
        except Exception:
            dett_perc = {}

        imp_tot = float(row["importo_totale"])
        riga = {
            "ID Spesa": row["id"],
            "Data Reg.": row["data_spesa"],
            "N° Fattura": row["num_fattura"],
            "Descrizione": row["descrizione"],
            "Rif. Pagamento Fattura": row.get("note_pagamento", ""),
            "Periodo Dal": row["periodo_dal"],
            "Periodo Al": row["periodo_al"],
            "Importo Totale (€)": imp_tot,
            "Commissione (€)": float(row["commissione"]),
        }
        for condomino in list_condomini:
            perc = float(dett_perc.get(condomino, 0.0))
            quota = (perc / 100.0) * imp_tot
            riga[f"{condomino} (%)"] = perc
            riga[f"{condomino} (€)"] = round(quota, 2)

        righe_excel.append(riga)

    df_excel = pd.DataFrame(righe_excel)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_excel.to_excel(writer, index=False, sheet_name="Teleriscaldamento")
    output.seek(0)
    return output


# ==========================================
# GESTIONE DATABASE SQLITE
# ==========================================
DB_FILE = "condominio_orchidea.db"


def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS condomini (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT UNIQUE
                )""")
    c.execute("""CREATE TABLE IF NOT EXISTS spese (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tipo TEXT,
                    descrizione TEXT,
                    num_fattura TEXT,
                    data_spesa TEXT,
                    periodo_dal TEXT,
                    periodo_al TEXT,
                    importo_totale REAL,
                    commissione REAL DEFAULT 0.0,
                    dettagli_json TEXT,
                    reportata INTEGER DEFAULT 0,
                    note_pagamento TEXT DEFAULT '',
                    num_condomini_divisore INTEGER DEFAULT 6,
                    esclusi_json TEXT DEFAULT '[]'
                )""")

    # Tabella incassi legata al singolo ID spesa
    c.execute("""CREATE TABLE IF NOT EXISTS incassi_spesa_condomino (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    spesa_id INTEGER,
                    condomino TEXT,
                    quota_dovuta REAL,
                    pagato INTEGER DEFAULT 0,
                    esente INTEGER DEFAULT 0,
                    note_incasso TEXT DEFAULT '',
                    UNIQUE(spesa_id, condomino)
                )""")

    c.execute("PRAGMA table_info(spese)")
    columns = [column[1] for column in c.fetchall()]
    if "commissione" not in columns:
        c.execute("ALTER TABLE spese ADD COLUMN commissione REAL DEFAULT 0.0")
    if "periodo_dal" not in columns:
        c.execute("ALTER TABLE spese ADD COLUMN periodo_dal TEXT DEFAULT ''")
    if "periodo_al" not in columns:
        c.execute("ALTER TABLE spese ADD COLUMN periodo_al TEXT DEFAULT ''")
    if "note_pagamento" not in columns:
        c.execute("ALTER TABLE spese ADD COLUMN note_pagamento TEXT DEFAULT ''")
    if "num_condomini_divisore" not in columns:
        c.execute(
            "ALTER TABLE spese ADD COLUMN num_condomini_divisore INTEGER DEFAULT 6"
        )
    if "esclusi_json" not in columns:
        c.execute("ALTER TABLE spese ADD COLUMN esclusi_json TEXT DEFAULT '[]'")

    c.execute("SELECT COUNT(*) FROM condomini")
    if c.fetchone()[0] == 0:
        default_condomini = [
            ("App. 1 - Rossi",),
            ("App. 2 - Bianchi",),
            ("App. 3 - Verdi",),
            ("App. 4 - Neri",),
            ("App. 5 - Ferrari",),
            ("App. 6 - Romano",),
        ]
        c.executemany("INSERT INTO condomini (nome) VALUES (?)", default_condomini)

    conn.commit()
    conn.close()


init_db()


def get_condomini():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT nome FROM condomini", conn)
    conn.close()
    return df["nome"].tolist()


# ==========================================
# INTERFACCIA STREAMLIT
# ==========================================
st.title("🌺 Gestione Spese & Archivio - Condominio Orchidea")

list_condomini = get_condomini()

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    [
        "📥 Carica Spesa / Fattura",
        "🗃️ Archivio & Modifica Spese",
        "🧾 Genera Report Condomino",
        "💳 Incassi per Spesa",
        "⚙️ Nomi Condomini",
    ]
)

# ==========================================
# TAB 1: CARICA SPESA / FATTURA
# ==========================================
with tab1:
    st.header("1. Inserisci una Nuova Spesa nell'Archivio")

    if "spesa_salvata" not in st.session_state:
        st.session_state["spesa_salvata"] = False

    if "form_reset_id" not in st.session_state:
        st.session_state["form_reset_id"] = 0

    reset_id = st.session_state["form_reset_id"]

    tipo_spesa = st.selectbox(
        "Seleziona Tipo di Spesa:",
        [
            "Teleriscaldamento",
            "Luce / Acqua / Utenze Comuni",
            "Spesa Varie (Scontrini/Pulizie/Giardino)",
        ],
        key=f"tipo_spesa_{reset_id}",
    )

    col1, col2 = st.columns(2)
    with col1:
        num_fattura = st.text_input(
            "N° Fattura / Ricevuta (opzionale)", value="", key=f"num_fattura_{reset_id}"
        )
        data_spesa = st.date_input(
            "Data di Registrazione/Emissione",
            datetime.now(),
            format="DD/MM/YYYY",
            key=f"data_spesa_{reset_id}",
        )

    with col2:
        descrizione_spesa = st.text_input(
            "Descrizione / Note",
            value="",
            placeholder="Es. Bolletta Luce Scale, Pulizia Giardino...",
            key=f"desc_spesa_{reset_id}",
        )
        importo_fattura = st.number_input(
            "Importo Fattura/Scontrino (€)",
            min_value=0.0,
            value=0.0,
            step=10.0,
            format="%.2f",
            key=f"imp_fattura_{reset_id}",
        )

    st.markdown("---")
    st.subheader("📅 Periodo di Consumo della Fattura / Spesa")
    col_per1, col_per2 = st.columns(2)
    with col_per1:
        periodo_dal = st.date_input(
            "Periodo Consumo - Dal:",
            datetime.now(),
            format="DD/MM/YYYY",
            key=f"per_dal_{reset_id}",
        )
    with col_per2:
        periodo_al = st.date_input(
            "Periodo Consumo - Al:",
            datetime.now(),
            format="DD/MM/YYYY",
            key=f"per_al_{reset_id}",
        )

    st.markdown("---")
    st.subheader("💳 Modalità di Pagamento e Commissione Bancaria")

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        metodo_pagamento = st.selectbox(
            "Seleziona Metodo di Pagamento:",
            [
                "Contanti / Gratuito (€ 0,00)",
                "Bonifico Bancario (+ € 1,00)",
                "PagoPA (+ € 1,00)",
                "Altra Commissione Personalizzata",
            ],
            key=f"metodo_pag_{reset_id}",
        )
        note_pagamento = st.text_input(
            "Riferimento Pagamento Fattura (CRO / TRN / PagoPA / Note)",
            value="",
            placeholder="Es. CRO: 1234567890123456",
            key=f"note_pag_{reset_id}",
        )

    with col_p2:
        if metodo_pagamento in ["Bonifico Bancario (+ € 1,00)", "PagoPA (+ € 1,00)"]:
            costo_commissione = 1.00
            st.info("Aggiunto automatico di **€ 1,00** per commissione di pagamento.")
        elif metodo_pagamento == "Altra Commissione Personalizzata":
            costo_commissione = st.number_input(
                "Inserisci Importo Commissione (€)",
                min_value=0.0,
                value=1.00,
                step=0.50,
                format="%.2f",
                key=f"comm_custom_{reset_id}",
            )
        else:
            costo_commissione = 0.00
            st.caption("Nessuna commissione applicata.")

    importo_totale_finale = importo_fattura + costo_commissione
    st.markdown(
        f"### 💰 **Importo Totale da Ripartire: € {importo_totale_finale:.2f}** *(Fattura € {importo_fattura:.2f} + Comm. € {costo_commissione:.2f})*"
    )

    st.markdown("---")
    dettagli_extra = {}
    esclusi_list = []
    divisore_calcolato = len(list_condomini)

    if tipo_spesa == "Teleriscaldamento":
        st.subheader("Percentuali di utilizzo per condomino (%):")
        df_tele = pd.DataFrame(
            {"Condomino": list_condomini, "Perc_Utilizzo": [0.0] * len(list_condomini)}
        )
        edited_tele = st.data_editor(
            df_tele,
            hide_index=True,
            use_container_width=True,
            key=f"editor_nuova_spesa_{reset_id}",
        )
        dettagli_extra = dict(
            zip(edited_tele["Condomino"], edited_tele["Perc_Utilizzo"])
        )
    else:
        st.subheader("👥 Ripartizione Condomini Spesa Generale")
        esclusi_list = st.multiselect(
            "Seleziona eventuali condomini ESENTI da questa spesa (es. alloggi vuoti/non interessati):",
            options=list_condomini,
            default=[],
            help="I condomini selezionati non pagheranno questa spesa. L'importo verrà diviso in parti uguali solo tra i rimanenti.",
            key=f"esclusi_{reset_id}",
        )
        attivi = [c for c in list_condomini if c not in esclusi_list]
        divisore_calcolato = len(attivi) if len(attivi) > 0 else 1
        st.info(
            f"💡 Spesa ripartita su **{divisore_calcolato}** condomini attivi. (Quota singola: **€ {importo_totale_finale / divisore_calcolato:.2f}**)"
        )

    if st.session_state["spesa_salvata"]:
        st.success("✅ Spesa salvata in archivio con successo!")
        if st.button("➕ Inserisci un'altra spesa", type="primary"):
            st.session_state["spesa_salvata"] = False
            st.session_state["form_reset_id"] += 1
            st.rerun()
    else:
        if st.button("💾 Salva Spesa in Archivio", type="primary"):
            if importo_fattura <= 0 and tipo_spesa != "Teleriscaldamento":
                st.warning(
                    "⚠️ Inserisci un importo maggiore di € 0.00 prima di salvare!"
                )
            elif not st.session_state["spesa_salvata"]:
                st.session_state["spesa_salvata"] = True
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute(
                    """INSERT INTO spese (tipo, descrizione, num_fattura, data_spesa, periodo_dal, periodo_al, importo_totale, commissione, dettagli_json, reportata, note_pagamento, num_condomini_divisore, esclusi_json)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?)""",
                    (
                        tipo_spesa,
                        descrizione_spesa if descrizione_spesa else tipo_spesa,
                        num_fattura,
                        data_spesa.strftime("%d/%m/%Y"),
                        periodo_dal.strftime("%d/%m/%Y"),
                        periodo_al.strftime("%d/%m/%Y"),
                        importo_totale_finale,
                        costo_commissione,
                        str(dettagli_extra),
                        note_pagamento,
                        divisore_calcolato,
                        str(esclusi_list),
                    ),
                )
                conn.commit()
                conn.close()
                st.rerun()

# ==========================================
# TAB 2: ARCHIVIO & MODIFICA SPESE
# ==========================================
with tab2:
    st.header("2. Archivio Generale e Modifica Spese Registrate")

    conn = sqlite3.connect(DB_FILE)
    df_spese = pd.read_sql_query(
        "SELECT id, tipo, descrizione, num_fattura, data_spesa, periodo_dal, periodo_al, importo_totale, commissione, dettagli_json, reportata, note_pagamento, num_condomini_divisore, esclusi_json FROM spese ORDER BY id DESC",
        conn,
    )
    conn.close()

    if not df_spese.empty:
        df_display = df_spese.copy()
        df_display["Data Reg."] = df_display["data_spesa"].apply(formatta_data_ita)
        df_display["Stato Report"] = df_display["reportata"].apply(
            lambda x: "✅ Già Inserita" if x == 1 else "⏳ Da Reportare"
        )
        df_display["Periodo Consumo"] = (
            df_display["periodo_dal"] + " - " + df_display["periodo_al"]
        )
        df_display["importo_fattura"] = (
            df_display["importo_totale"] - df_display["commissione"]
        )

        st.dataframe(
            df_display[
                [
                    "id",
                    "Data Reg.",
                    "tipo",
                    "descrizione",
                    "num_fattura",
                    "Periodo Consumo",
                    "importo_fattura",
                    "commissione",
                    "importo_totale",
                    "num_condomini_divisore",
                    "note_pagamento",
                    "Stato Report",
                ]
            ],
            use_container_width=True,
            column_config={
                "importo_fattura": st.column_config.NumberColumn(
                    "Importo Fattura (€)", format="€ %.2f"
                ),
                "commissione": st.column_config.NumberColumn(
                    "Comm. (€)", format="€ %.2f"
                ),
                "importo_totale": st.column_config.NumberColumn(
                    "Totale Ripartito (€)", format="€ %.2f"
                ),
                "num_condomini_divisore": st.column_config.NumberColumn(
                    "N° Divisore", format="%d"
                ),
                "note_pagamento": st.column_config.TextColumn("Rif. Pagamento Fattura"),
            },
        )

        df_tele_only = df_spese[df_spese["tipo"] == "Teleriscaldamento"]
        if not df_tele_only.empty:
            st.markdown("---")
            st.subheader("📊 Report Ripartizione Teleriscaldamento")
            excel_bytes = genera_excel_teleriscaldamento(df_tele_only, list_condomini)
            st.download_button(
                label="📊 Scarica Report Teleriscaldamento (Excel .xlsx)",
                data=excel_bytes,
                file_name=f"Riepilogo_Teleriscaldamento_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        st.markdown("---")
        st.subheader("✏️ Modifica o Elimina una Spesa Esistente")

        opzioni_spese = {
            f"ID {row['id']} - {row['descrizione']} (€ {row['importo_totale']:.2f})": row[
                "id"
            ]
            for _, row in df_spese.iterrows()
        }
        spesa_selezionata_label = st.selectbox(
            "Seleziona la spesa da modificare/eliminare:", list(opzioni_spese.keys())
        )

        id_spesa_sel = opzioni_spese[spesa_selezionata_label]
        spesa_dati = df_spese[df_spese["id"] == id_spesa_sel].iloc[0]

        with st.expander("📝 Apri Modulo di Modifica Spesa", expanded=True):
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                mod_tipo = st.selectbox(
                    "Tipo Spesa",
                    [
                        "Teleriscaldamento",
                        "Luce / Acqua / Utenze Comuni",
                        "Spesa Varie (Scontrini/Pulizie/Giardino)",
                    ],
                    index=(
                        [
                            "Teleriscaldamento",
                            "Luce / Acqua / Utenze Comuni",
                            "Spesa Varie (Scontrini/Pulizie/Giardino)",
                        ].index(spesa_dati["tipo"])
                        if spesa_dati["tipo"]
                        in [
                            "Teleriscaldamento",
                            "Luce / Acqua / Utenze Comuni",
                            "Spesa Varie (Scontrini/Pulizie/Giardino)",
                        ]
                        else 0
                    ),
                    key=f"mod_tipo_{id_spesa_sel}",
                )
                mod_num_fattura = st.text_input(
                    "N° Fattura",
                    value=str(spesa_dati["num_fattura"] or ""),
                    key=f"mod_num_fatt_{id_spesa_sel}",
                )
                mod_data_spesa = st.date_input(
                    "Data Emissione",
                    value=parse_data_ita(spesa_dati["data_spesa"]),
                    format="DD/MM/YYYY",
                    key=f"mod_data_{id_spesa_sel}",
                )
                mod_note_pagamento = st.text_input(
                    "Riferimento Pagamento Fattura (CRO/TRN/PagoPA)",
                    value=str(spesa_dati.get("note_pagamento", "") or ""),
                    key=f"mod_note_pag_{id_spesa_sel}",
                )

            with col_m2:
                mod_descrizione = st.text_input(
                    "Descrizione",
                    value=str(spesa_dati["descrizione"] or ""),
                    key=f"mod_desc_{id_spesa_sel}",
                )
                imp_base = float(spesa_dati["importo_totale"]) - float(
                    spesa_dati["commissione"]
                )
                mod_importo_base = st.number_input(
                    "Importo Fattura/Scontrino (€)",
                    min_value=0.0,
                    value=max(0.0, imp_base),
                    step=10.0,
                    format="%.2f",
                    key=f"mod_imp_{id_spesa_sel}",
                )
                mod_commissione = st.number_input(
                    "Commissione (€)",
                    min_value=0.0,
                    value=float(spesa_dati["commissione"]),
                    step=0.50,
                    format="%.2f",
                    key=f"mod_comm_{id_spesa_sel}",
                )

            col_mp1, col_mp2 = st.columns(2)
            with col_mp1:
                mod_periodo_dal = st.date_input(
                    "Periodo Dal",
                    value=parse_data_ita(spesa_dati["periodo_dal"]),
                    format="DD/MM/YYYY",
                    key=f"mod_dal_{id_spesa_sel}",
                )
            with col_mp2:
                mod_periodo_al = st.date_input(
                    "Periodo Al",
                    value=parse_data_ita(spesa_dati["periodo_al"]),
                    format="DD/MM/YYYY",
                    key=f"mod_al_{id_spesa_sel}",
                )
            mod_totale = mod_importo_base + mod_commissione

            mod_dettagli_extra = {}
            mod_esclusi_list = []
            mod_divisore = len(list_condomini)

            if mod_tipo == "Teleriscaldamento":
                st.subheader("Percentuali di utilizzo per condomino (%):")
                dett_salvati = {}
                try:
                    if spesa_dati["dettagli_json"]:
                        dett_salvati = ast.literal_eval(spesa_dati["dettagli_json"])
                except Exception:
                    dett_salvati = {}

                valori_perc = [float(dett_salvati.get(c, 0.0)) for c in list_condomini]
                df_tele_mod = pd.DataFrame(
                    {"Condomino": list_condomini, "Perc_Utilizzo": valori_perc}
                )
                edited_tele_mod = st.data_editor(
                    df_tele_mod,
                    hide_index=True,
                    use_container_width=True,
                    key=f"editor_mod_{id_spesa_sel}",
                )
                mod_dettagli_extra = dict(
                    zip(edited_tele_mod["Condomino"], edited_tele_mod["Perc_Utilizzo"])
                )
            else:
                st.subheader("👥 Ripartizione Condomini Spesa Generale")
                esclusi_salvati = []
                try:
                    if spesa_dati["esclusi_json"]:
                        esclusi_salvati = ast.literal_eval(spesa_dati["esclusi_json"])
                except Exception:
                    esclusi_salvati = []

                mod_esclusi_list = st.multiselect(
                    "Seleziona condomini ESENTI da questa spesa:",
                    options=list_condomini,
                    default=[c for c in esclusi_salvati if c in list_condomini],
                    key=f"mod_esclusi_{id_spesa_sel}",
                )
                attivi_mod = [c for c in list_condomini if c not in mod_esclusi_list]
                mod_divisore = len(attivi_mod) if len(attivi_mod) > 0 else 1
                st.info(f"💡 Spesa divisa su **{mod_divisore}** condomini attivi.")

            st.markdown("---")
            conferma_modifica = st.checkbox(
                "⚠️ Confermo di voler sovrascrivere questa spesa con i nuovi dati",
                key=f"check_conf_{id_spesa_sel}",
            )

            col_btn_m1, col_btn_m2 = st.columns(2)
            with col_btn_m1:
                if st.button("💾 Salva Modifiche", type="primary", key="btn_salva_mod"):
                    if conferma_modifica:
                        conn = sqlite3.connect(DB_FILE)
                        c = conn.cursor()
                        c.execute(
                            """UPDATE spese 
                                     SET tipo = ?, descrizione = ?, num_fattura = ?, data_spesa = ?, 
                                         periodo_dal = ?, periodo_al = ?, importo_totale = ?, commissione = ?, dettagli_json = ?, note_pagamento = ?, num_condomini_divisore = ?, esclusi_json = ?
                                     WHERE id = ?""",
                            (
                                mod_tipo,
                                mod_descrizione,
                                mod_num_fattura,
                                mod_data_spesa.strftime("%d/%m/%Y"),
                                mod_periodo_dal.strftime("%d/%m/%Y"),
                                mod_periodo_al.strftime("%d/%m/%Y"),
                                mod_totale,
                                mod_commissione,
                                str(mod_dettagli_extra),
                                mod_note_pagamento,
                                mod_divisore,
                                str(mod_esclusi_list),
                                id_spesa_sel,
                            ),
                        )
                        conn.commit()
                        conn.close()
                        st.success("✅ Spesa aggiornata con successo!")
                        st.rerun()
                    else:
                        st.warning(
                            "⚠️ Per salvare spunta prima la casella di conferma!"
                        )

            with col_btn_m2:
                if st.button(
                    "🗑️ Elimina Questa Spesa", type="secondary", key="btn_elimina_spesa"
                ):
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    c.execute("DELETE FROM spese WHERE id = ?", (id_spesa_sel,))
                    c.execute(
                        "DELETE FROM incassi_spesa_condomino WHERE spesa_id = ?",
                        (id_spesa_sel,),
                    )
                    conn.commit()
                    conn.close()
                    st.warning("🗑️ Spesa eliminata!")
                    st.rerun()

    else:
        st.info("Nessuna spesa salvata in archivio.")

# ==========================================
# TAB 3: GENERAZIONE REPORT PERSONALIZZATO CONDOMINO
# ==========================================
with tab3:
    st.header("3. 📑 Genera Report Personalizzato e Stampa Avviso Condomino")
    st.info(
        "Seleziona le spese/fatture che desideri addebitare in questo report. Il sistema calcolerà in automatico l'importo per ogni singola fattura in base alle ripartizioni salvate e genererà il PDF finale con il totale."
    )

    conn = sqlite3.connect(DB_FILE)
    df_filtrato = pd.read_sql_query("SELECT * FROM spese ORDER BY id DESC", conn)
    conn.close()

    if not df_filtrato.empty:
        st.subheader("1. Seleziona le Spese/Fatture da Includere nel Report:")

        # Di default suggerisce le spese non ancora inserite in altri report
        df_filtrato["Includi"] = df_filtrato["reportata"] == 0

        df_filtrato["Data Reg."] = df_filtrato["data_spesa"].apply(formatta_data_ita)
        df_filtrato["Stato"] = df_filtrato["reportata"].apply(
            lambda x: "⚠️ GIÀ REPORTATA" if x == 1 else "🟢 NUOVA"
        )
        df_filtrato["Periodo"] = (
            df_filtrato["periodo_dal"] + " - " + df_filtrato["periodo_al"]
        )

        edited_spese_rep = st.data_editor(
            df_filtrato[
                [
                    "Includi",
                    "id",
                    "Data Reg.",
                    "tipo",
                    "descrizione",
                    "num_fattura",
                    "Periodo",
                    "importo_totale",
                    "Stato",
                ]
            ],
            use_container_width=True,
            column_config={
                "num_fattura": st.column_config.TextColumn("N° Fattura"),
                "importo_totale": st.column_config.NumberColumn(
                    "Totale Spesa (€)", format="€ %.2f"
                ),
            },
            disabled=[
                "id",
                "Data Reg.",
                "tipo",
                "descrizione",
                "num_fattura",
                "Periodo",
                "importo_totale",
                "Stato",
            ],
            hide_index=True,
            key="editor_report_personalizzato",
        )

        spese_selezionate_ids = edited_spese_rep[edited_spese_rep["Includi"] == True][
            "id"
        ].tolist()
        spese_scelte_df = df_filtrato[df_filtrato["id"].isin(spese_selezionate_ids)]

        st.markdown("---")
        st.subheader("2. Impostazioni Intestazione e Condomino")

        lista_descrizioni = (
            spese_scelte_df["descrizione"].unique().tolist()
            if not spese_scelte_df.empty
            else []
        )
        intestazione_automatica = (
            " - ".join(lista_descrizioni)
            if lista_descrizioni
            else "Prospetto Spese Condominiali"
        )

        col_c1, col_c2 = st.columns(2)
        with col_c1:
            condomino_sel = (
                st.selectbox("Seleziona Condomino per la Stampa PDF:", list_condomini)
                if list_condomini
                else None
            )
            periodo_testo = st.text_input(
                "Oggetto / Oggetto dell'Avviso:", value=intestazione_automatica
            )
        with col_c2:
            mostra_scad = st.checkbox(
                "Inserisci Data di Scadenza Pagamento?", value=True
            )
            scad_val = (
                st.date_input(
                    "Data Scadenza", datetime.now(), format="DD/MM/YYYY"
                ).strftime("%d/%m/%Y")
                if mostra_scad
                else ""
            )

        totale_condomino = 0.0
        dettagli_pdf = []

        if condomino_sel and not spese_scelte_df.empty:
            for idx, row in spese_scelte_df.iterrows():
                imp = row["importo_totale"]
                t_spesa = row["tipo"]
                comm = row.get("commissione", 0.0)
                n_fatt = row.get("num_fattura", "")
                p_dal = row.get("periodo_dal", "")
                p_al = row.get("periodo_al", "")
                divisore_spesa = int(row.get("num_condomini_divisore", 6) or 6)

                esclusi_spesa = []
                try:
                    if row.get("esclusi_json"):
                        esclusi_spesa = ast.literal_eval(row["esclusi_json"])
                except Exception:
                    esclusi_spesa = []

                str_periodo = f" (Dal {p_dal} al {p_al})" if p_dal and p_al else ""
                str_fattura = f" - Fattura N° {n_fatt}" if n_fatt else ""

                note_desc = f"{row['descrizione']}{str_fattura}{str_periodo}"
                if comm > 0:
                    note_desc += f" [incl. comm. EUR {comm:.2f}]"

                if t_spesa == "Teleriscaldamento":
                    dett_perc = {}
                    try:
                        if row["dettagli_json"]:
                            dett_perc = ast.literal_eval(row["dettagli_json"])
                    except Exception:
                        dett_perc = {}
                    perc = float(dett_perc.get(condomino_sel, 0.0))

                    q = (perc / 100.0) * imp
                    q_arrotondato = round(
                        q, 2
                    )  # <--- MODIFICA: Arrotondiamo subito a 2 decimali

                    dettagli_pdf.append(
                        (
                            pulisci_testo_pdf(note_desc),
                            f"Consumi ({perc:.1f}%)",
                            q_arrotondato,
                        )
                    )
                    totale_condomino += (
                        q_arrotondato  # <--- MODIFICA: Sommiamo l'importo arrotondato
                    )
                else:
                    if condomino_sel in esclusi_spesa:
                        dettagli_pdf.append(
                            (
                                pulisci_testo_pdf(note_desc),
                                "Esente (Quota 0)",
                                0.0,
                            )
                        )
                    else:
                        q = imp / divisore_spesa if divisore_spesa > 0 else 0
                        q_arrotondato = round(
                            q, 2
                        )  # <--- MODIFICA: Arrotondiamo subito a 2 decimali

                        dettagli_pdf.append(
                            (
                                pulisci_testo_pdf(note_desc),
                                f"Quota 1/{divisore_spesa}",
                                q_arrotondato,
                            )
                        )
                        totale_condomino += q_arrotondato  # <--- MODIFICA: Sommiamo l'importo arrotondato

            st.markdown("---")
            st.markdown(
                f"### 💰 Totale Calcolato per **{condomino_sel}**: **€ {totale_condomino:.2f}**"
            )

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("👁️ Genera Report PDF Stampabile", type="primary"):

                    class PDFRicevuta(FPDF):
                        def header(self):
                            self.set_font("Helvetica", "B", 18)
                            self.set_text_color(44, 62, 80)
                            self.cell(0, 10, "CONDOMINIO ORCHIDEA", ln=True)
                            self.set_font("Helvetica", "", 10)
                            self.set_text_color(127, 140, 141)
                            self.cell(
                                0,
                                5,
                                "Prospetto Ripartizione Spese e Avviso di Pagamento",
                                ln=True,
                            )
                            self.ln(5)
                            self.line(10, self.get_y(), 200, self.get_y())
                            self.ln(8)

                    pdf = PDFRicevuta()
                    pdf.add_page()
                    pdf.set_font("Helvetica", "B", 12)
                    pdf.cell(
                        0,
                        6,
                        pulisci_testo_pdf(f"Destinatario: {condomino_sel}"),
                        ln=True,
                    )
                    pdf.set_font("Helvetica", "", 10)
                    pdf.cell(
                        0,
                        5,
                        pulisci_testo_pdf(
                            f"Oggetto: {periodo_testo} | Data emissione: {datetime.now().strftime('%d/%m/%Y')}"
                        ),
                        ln=True,
                    )
                    pdf.ln(8)

                    # Tabella intestazione
                    pdf.set_fill_color(44, 62, 80)
                    pdf.set_text_color(255, 255, 255)
                    pdf.set_font("Helvetica", "B", 9.5)
                    pdf.cell(105, 8, " Dettaglio Spesa / Fattura / Periodo", fill=True)
                    pdf.cell(45, 8, " Criterio", fill=True)
                    pdf.cell(
                        40, 8, " Quota Spettante (EUR)", fill=True, ln=True, align="R"
                    )

                    pdf.set_text_color(50, 50, 50)
                    pdf.set_font("Helvetica", "", 9)

                    # Ciclo per singola fattura selezionata
                    for item in dettagli_pdf:
                        y_inizio = pdf.get_y()
                        pdf.set_xy(10, y_inizio)
                        pdf.multi_cell(105, 5, item[0], border=0)
                        y_col1 = pdf.get_y()

                        pdf.set_xy(115, y_inizio)
                        pdf.multi_cell(45, 5, item[1], border=0)
                        y_col2 = pdf.get_y()

                        pdf.set_xy(160, y_inizio)
                        pdf.cell(40, 5, f"{item[2]:.2f}", border=0, align="R")
                        y_col3 = pdf.get_y() + 5

                        y_max = max(y_col1, y_col2, y_col3)
                        pdf.set_y(y_max)
                        pdf.set_draw_color(220, 220, 220)
                        pdf.line(10, y_max, 200, y_max)
                        pdf.set_y(y_max + 3)

                    # Totale in fondo al PDF
                    pdf.ln(8)
                    pdf.set_fill_color(234, 242, 248)
                    pdf.rect(110, pdf.get_y(), 90, 22, style="F")
                    pdf.set_xy(115, pdf.get_y() + 3)
                    pdf.set_font("Helvetica", "B", 10)
                    pdf.set_text_color(44, 62, 80)
                    pdf.cell(80, 5, "TOTALE IMPORTO DOVUTO:", align="R", ln=True)
                    pdf.set_x(115)
                    pdf.set_font("Helvetica", "B", 15)
                    pdf.set_text_color(26, 82, 118)
                    pdf.cell(80, 7, f"EUR {totale_condomino:.2f}", align="R", ln=True)

                    if mostra_scad:
                        pdf.ln(8)
                        pdf.set_font("Helvetica", "B", 10)
                        pdf.set_text_color(192, 57, 43)
                        pdf.cell(
                            0, 5, f"Scadenza pagamento: {scad_val}", align="R", ln=True
                        )

                    pdf_output_bytes = bytes(pdf.output())
                    st.session_state["pdf_pronto"] = pdf_output_bytes
                    st.session_state["pdf_nome_file"] = (
                        f"Report_{condomino_sel.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf"
                    )

            with col_btn2:
                if st.button("🔒 Segna le spese selezionate come 'GIÀ REPORTATE'"):
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    for sp_id in spese_selezionate_ids:
                        c.execute(
                            "UPDATE spese SET reportata = 1 WHERE id = ?", (sp_id,)
                        )
                    conn.commit()
                    conn.close()
                    st.success("✅ Spese aggiornate come 'Già Reportate'!")
                    st.rerun()

            if "pdf_pronto" in st.session_state and st.session_state["pdf_pronto"]:
                st.markdown("---")
                st.subheader("📄 Anteprima del Report PDF Personalizzato")
                mostra_pdf_in_anteprima(st.session_state["pdf_pronto"])
                st.download_button(
                    label="💾 Scarica e Stampa PDF",
                    data=st.session_state["pdf_pronto"],
                    file_name=st.session_state["pdf_nome_file"],
                    mime="application/pdf",
                    type="primary",
                )
        elif spese_scelte_df.empty:
            st.warning(
                "⚠️ Seleziona almeno una spesa/fattura dalla tabella in alto per generare il report."
            )
    else:
        st.info("Nessuna spesa trovata in archivio.")

# ==========================================
# TAB 4: GESTIONE INCASSI PER SINGOLA SPESA E STORICO CONDOMINI
# ==========================================
with tab4:
    st.header("💳 Gestione Incassi per Singola Spesa e Storico Condomini")

    conn = sqlite3.connect(DB_FILE)
    df_spese_all = pd.read_sql_query("SELECT * FROM spese ORDER BY id DESC", conn)
    df_incassi_all = pd.read_sql_query("SELECT * FROM incassi_spesa_condomino", conn)
    conn.close()

    if not df_spese_all.empty:
        # Calcoliamo lo stato di saldo di ogni spesa (se tutti i condomini attivi hanno saldato)
        spese_da_saldare_ids = []
        map_spesa_completata = {}

        for _, sp in df_spese_all.iterrows():
            sp_id = sp["id"]
            incassi_sp = df_incassi_all[df_incassi_all["spesa_id"] == sp_id]

            # Se ci sono record incasso per tutti i condomini e sono tutti pagati/esenti
            if not incassi_sp.empty and len(incassi_sp) >= len(list_condomini):
                tutti_pagati = all(
                    (row["pagato"] == 1 or row["esente"] == 1)
                    for _, row in incassi_sp.iterrows()
                )
                map_spesa_completata[sp_id] = tutti_pagati
            else:
                map_spesa_completata[sp_id] = False

            if not map_spesa_completata[sp_id]:
                spese_da_saldare_ids.append(sp_id)

        col_f1, col_f2 = st.columns([2, 1])
        with col_f2:
            mostra_tutte = st.checkbox(
                "👁️ Mostra anche le spese GIÀ SALDATE da tutti",
                value=False,
                help="Di default vengono mostrate solo le spese che hanno ancora incassi in sospeso.",
            )

        # Filtriamo le spese in base al pulsante
        if mostra_tutte:
            spese_visibili_df = df_spese_all
        else:
            spese_visibili_df = df_spese_all[
                df_spese_all["id"].isin(spese_da_saldare_ids)
            ]
            if spese_visibili_df.empty:
                st.success(
                    "🎉 Tutte le spese in archivio risultano completamente saldate!"
                )
                spese_visibili_df = (
                    df_spese_all  # Fallback per mostrare comunque la lista
                )

        # Mappatura opzioni per la SelectBox
        opzioni_spese_dict = {}
        for _, sp in spese_visibili_df.iterrows():
            stato_icona = (
                "✅ SALDATA"
                if map_spesa_completata.get(sp["id"], False)
                else "⏳ IN INCASSO"
            )
            n_fatt_str = f" Fatt. {sp['num_fattura']}" if sp["num_fattura"] else ""
            label = f"ID {sp['id']} - {sp['descrizione']}{n_fatt_str} (€ {sp['importo_totale']:.2f}) [{sp['data_spesa']}] - {stato_icona}"
            opzioni_spese_dict[label] = sp["id"]

        st.subheader("1. Seleziona la Spesa da Incassare")
        spesa_incasso_label = st.selectbox(
            "Seleziona Spesa:",
            list(opzioni_spese_dict.keys()),
            key="select_spesa_incasso",
        )

        spesa_incasso_id = opzioni_spese_dict[spesa_incasso_label]
        spesa_obj = df_spese_all[df_spese_all["id"] == spesa_incasso_id].iloc[0]

        # Calcolo quote teoriche per la spesa selezionata
        imp_spesa = float(spesa_obj["importo_totale"])
        tipo_sp = spesa_obj["tipo"]
        divisore_sp = int(spesa_obj.get("num_condomini_divisore", 6) or 6)

        esclusi_sp = []
        try:
            if spesa_obj.get("esclusi_json"):
                esclusi_sp = ast.literal_eval(spesa_obj["esclusi_json"])
        except Exception:
            esclusi_sp = []

        dett_tele_p = {}
        if tipo_sp == "Teleriscaldamento":
            try:
                if spesa_obj.get("dettagli_json"):
                    dett_tele_p = ast.literal_eval(spesa_obj["dettagli_json"])
            except Exception:
                pass

        # Recuperiamo gli incassi già salvati per questa spesa
        df_incassi_spesa = df_incassi_all[
            df_incassi_all["spesa_id"] == spesa_incasso_id
        ]

        st.markdown(
            f"#### 📄 Incassi per: **{spesa_obj['descrizione']}** (Importo Totale: **€ {imp_spesa:.2f}**)"
        )
        st.markdown("---")

        col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns([2.5, 1, 1.2, 1, 3.5])
        col_h1.markdown("**Condomino**")
        col_h2.markdown("**Esente?**")
        col_h3.markdown("**Quota Dovuta**")
        col_h4.markdown("**Saldato?**")
        col_h5.markdown("**Rif. Bonifico / Data Versamento**")

        nuovi_dati_spesa_incassata = []
        tot_dovuto_spesa = 0.0
        tot_incassato_spesa = 0.0

        for idx, c_nome in enumerate(list_condomini):
            # Calcolo quota condomino per questa spesa
            if tipo_sp == "Teleriscaldamento":
                p = float(dett_tele_p.get(c_nome, 0.0))
                q_calc = (p / 100.0) * imp_spesa
            else:
                q_calc = (
                    0.0
                    if c_nome in esclusi_sp
                    else (
                        imp_spesa / divisore_sp_if_valid
                        if (
                            divisore_sp_if_valid := (
                                divisore_sp if divisore_sp > 0 else 1
                            )
                        )
                        else 0
                    )
                )

            # Recupero dati da DB se esistono
            r_db = (
                df_incassi_spesa[df_incassi_spesa["condomino"] == c_nome]
                if not df_incassi_spesa.empty
                else pd.DataFrame()
            )

            is_pag = bool(r_db.iloc[0]["pagato"]) if not r_db.empty else False
            is_ese = (
                bool(r_db.iloc[0]["esente"])
                if not r_db.empty
                else (c_nome in esclusi_sp or q_calc == 0.0)
            )
            note_val_db = (
                str(r_db.iloc[0]["note_incasso"] or "") if not r_db.empty else ""
            )

            col_r1, col_r2, col_r3, col_r4, col_r5 = st.columns([2.5, 1, 1.2, 1, 3.5])

            col_r1.write(f"**{c_nome}**")
            esente_val = col_r2.checkbox(
                "", value=is_ese, key=f"sp_ese_{spesa_incasso_id}_{idx}"
            )

            quota_fin = 0.00 if esente_val else round(q_calc, 2)
            col_r3.write(f"€ {quota_fin:.2f}")

            saldato_val = col_r4.checkbox(
                "", value=is_pag, key=f"sp_sal_{spesa_incasso_id}_{idx}"
            )
            note_val = col_r5.text_input(
                "",
                value=note_val_db,
                key=f"sp_not_{spesa_incasso_id}_{idx}",
                label_visibility="collapsed",
            )

            nuovi_dati_spesa_incassata.append(
                {
                    "condomino": c_nome,
                    "quota": quota_fin,
                    "esente": esente_val,
                    "saldato": saldato_val,
                    "note": note_val,
                }
            )

            tot_dovuto_spesa += quota_fin
            if saldato_val:
                tot_incassato_spesa += quota_fin

        tot_rimanente_spesa = tot_dovuto_spesa - tot_incassato_spesa

        st.markdown("---")
        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric("Totale Spesa", f"€ {tot_dovuto_spesa:.2f}")
        c_m2.metric("Incassato per questa spesa", f"€ {tot_incassato_spesa:.2f}")
        c_m3.metric(
            "Rimanente da Incassare",
            f"€ {tot_rimanente_spesa:.2f}",
            delta=-tot_rimanente_spesa if tot_rimanente_spesa > 0 else 0,
        )

        if st.button(
            "💾 Salva Incassi per Questa Spesa",
            type="primary",
            key="btn_salva_incasso_spesa",
        ):
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            for item in nuovi_dati_spesa_incassata:
                c.execute(
                    """INSERT INTO incassi_spesa_condomino (spesa_id, condomino, quota_dovuta, pagato, esente, note_incasso)
                             VALUES (?, ?, ?, ?, ?, ?)
                             ON CONFLICT(spesa_id, condomino) DO UPDATE SET
                                quota_dovuta=excluded.quota_dovuta,
                                pagato=excluded.pagato,
                                esente=excluded.esente,
                                note_incasso=excluded.note_incasso""",
                    (
                        spesa_incasso_id,
                        item["condomino"],
                        item["quota"],
                        1 if item["saldato"] else 0,
                        1 if item["esente"] else 0,
                        item["note"],
                    ),
                )
            conn.commit()
            conn.close()
            st.success(f"✅ Stato incassi per spesa ID {spesa_incasso_id} salvato!")
            st.rerun()

        # ==========================================
        # SEZIONE PUNTO 1: VISTA STORICO E SITUAZIONE SINGOLO CONDOMINO
        # ==========================================
        st.markdown("---")
        st.subheader("2. 📊 Vista Storico Pagamenti per Singolo Condomino")

        condomino_storico_sel = st.selectbox(
            "Seleziona un Condomino per vedere tutto la sua situazione storica:",
            list_condomini,
            key="select_condomino_storico",
        )

        if condomino_storico_sel:
            righe_storico_condomino = []
            tot_dovuto_cond = 0.0
            tot_versato_cond = 0.0

            for _, sp in df_spese_all.iterrows():
                s_id = sp["id"]
                imp_s = float(sp["importo_totale"])
                t_s = sp["tipo"]
                div_s = int(sp.get("num_condomini_divisore", 6) or 6)

                escl_s = []
                try:
                    if sp.get("esclusi_json"):
                        escl_s = ast.literal_eval(sp["esclusi_json"])
                except Exception:
                    escl_s = []

                if t_s == "Teleriscaldamento":
                    d_t = {}
                    try:
                        if sp.get("dettagli_json"):
                            d_t = ast.literal_eval(sp["dettagli_json"])
                    except Exception:
                        pass
                    p = float(d_t.get(condomino_storico_sel, 0.0))
                    q_c = (p / 100.0) * imp_s
                else:
                    q_c = (
                        0.0
                        if condomino_storico_sel in escl_s
                        else (imp_s / (div_s if div_s > 0 else 1))
                    )

                # Stato dal DB
                r_inc = df_incassi_all[
                    (df_incassi_all["spesa_id"] == s_id)
                    & (df_incassi_all["condomino"] == condomino_storico_sel)
                ]
                is_p = bool(r_inc.iloc[0]["pagato"]) if not r_inc.empty else False
                is_e = (
                    bool(r_inc.iloc[0]["esente"])
                    if not r_inc.empty
                    else (condomino_storico_sel in escl_s or q_c == 0.0)
                )
                n_i = (
                    str(r_inc.iloc[0]["note_incasso"] or "") if not r_inc.empty else ""
                )

                quota_effettiva = 0.00 if is_e else round(q_c, 2)

                righe_storico_condomino.append(
                    {
                        "ID Spesa": s_id,
                        "Data Emissione": sp["data_spesa"],
                        "Descrizione Spesa": sp["descrizione"],
                        "N° Fattura": sp["num_fattura"],
                        "Quota Dovuta (€)": quota_effettiva,
                        "Stato Pagamento": (
                            "🟢 SALDATO"
                            if is_p
                            else ("⚪ ESENTE" if is_e else "🔴 DA SALDARE")
                        ),
                        "Rif. / Note Versamento": n_i,
                    }
                )

                if not is_e:
                    tot_dovuto_cond += quota_effettiva
                    if is_p:
                        tot_versato_cond += quota_effettiva

            df_storico_cond = pd.DataFrame(righe_storico_condomino)
            tot_da_versare_cond = tot_dovuto_cond - tot_versato_cond

            st.markdown(f"#### Situazione Contabile per **{condomino_storico_sel}**")

            st.dataframe(
                df_storico_cond,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Quota Dovuta (€)": st.column_config.NumberColumn(
                        "Quota Dovuta (€)", format="€ %.2f"
                    )
                },
            )

            st.info(
                f"💰 Totale Dovuto Storico: **€ {tot_dovuto_cond:.2f}** | Totale Già Versato: **€ {tot_versato_cond:.2f}** | Rimanente da Pagare: **€ {tot_da_versare_cond:.2f}**"
            )

    else:
        st.info("Nessuna spesa trovata in archivio.")

# ==========================================
# TAB 5: NOMI CONDOMINI
# ==========================================
with tab5:
    st.header("Anagrafica Nomi Condomini")
    st.info("Modifica i nomi esistenti o aggiungi un nuovo condomino.")

    n_condomini = get_condomini()
    nuovi_nomi = []
    for i, nome in enumerate(n_condomini):
        nuovi_nomi.append(st.text_input(f"Condomino {i+1}", nome, key=f"condomino_{i}"))

    if st.button("💾 Aggiorna Nomi Condomini"):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("DELETE FROM condomini")
        for n in nuovi_nomi:
            if n.strip():
                c.execute("INSERT INTO condomini (nome) VALUES (?)", (n.strip(),))
        conn.commit()
        conn.close()
        st.success("Nomi aggiornati!")
        st.rerun()

    st.markdown("---")
    st.subheader("➕ Aggiungi un Nuovo Condomino")
    nuovo_condomino_nome = st.text_input("Nome e Cognome / App. nuovo condomino:")
    if st.button("➕ Aggiungi"):
        if nuovo_condomino_nome.strip():
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            try:
                c.execute(
                    "INSERT INTO condomini (nome) VALUES (?)",
                    (nuovo_condomino_nome.strip(),),
                )
                conn.commit()
                st.success(f"Condomino '{nuovo_condomino_nome}' aggiunto!")
            except sqlite3.IntegrityError:
                st.error("Esiste già un condomino con questo nome!")
            conn.close()
            st.rerun()
