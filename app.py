import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
from fpdf import FPDF
import ast
import base64
import io

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

    c.execute("""CREATE TABLE IF NOT EXISTS incassi_condomini (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    condomino TEXT,
                    periodo_mese_anno TEXT,
                    quota_dovuta REAL,
                    pagato INTEGER DEFAULT 0,
                    esente INTEGER DEFAULT 0,
                    note_incasso TEXT DEFAULT '',
                    UNIQUE(condomino, periodo_mese_anno)
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

    c.execute("PRAGMA table_info(incassi_condomini)")
    inc_columns = [column[1] for column in c.fetchall()]
    if "esente" not in inc_columns:
        c.execute("ALTER TABLE incassi_condomini ADD COLUMN esente INTEGER DEFAULT 0")

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
        "💳 Incassi Condomini",
        "⚙️ Nomi Condomini",
    ]
)

# ==========================================
# TAB 1: CARICA SPESA / FATTURA
# ==========================================
with tab1:
    st.header("1. Inserisci una Nuova Spesa nell'Archivio")

    tipo_spesa = st.selectbox(
        "Seleziona Tipo di Spesa:",
        [
            "Teleriscaldamento",
            "Luce / Acqua / Utenze Comuni",
            "Spesa Varie (Scontrini/Pulizie/Giardino)",
        ],
    )

    col1, col2 = st.columns(2)
    with col1:
        num_fattura = st.text_input(
            "N° Fattura / Ricevuta (opzionale)",
            value="" if "Spesa Varie" in tipo_spesa else "FAT-001",
        )
        data_spesa = st.date_input(
            "Data di Registrazione/Emissione", datetime.now(), format="DD/MM/YYYY"
        )

    with col2:
        descrizione_spesa = st.text_input("Descrizione / Note", value=tipo_spesa)
        importo_fattura = st.number_input(
            "Importo Fattura/Scontrino (€)",
            min_value=0.0,
            value=100.0,
            step=10.0,
            format="%.2f",
        )

    st.markdown("---")
    st.subheader("📅 Periodo di Consumo della Fattura / Spesa")
    col_per1, col_per2 = st.columns(2)
    with col_per1:
        periodo_dal = st.date_input(
            "Periodo Consumo - Dal:", datetime.now(), format="DD/MM/YYYY"
        )
    with col_per2:
        periodo_al = st.date_input(
            "Periodo Consumo - Al:", datetime.now(), format="DD/MM/YYYY"
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
        )
        note_pagamento = st.text_input(
            "Riferimento Pagamento Fattura (CRO / TRN / PagoPA / Note)",
            value="",
            help="Es. CRO: 1234567890123456",
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
            df_tele, hide_index=True, use_container_width=True, key="editor_nuova_spesa"
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
        )
        attivi = [c for c in list_condomini if c not in esclusi_list]
        divisore_calcolato = len(attivi) if len(attivi) > 0 else 1
        st.info(
            f"💡 Spesa ripartita su **{divisore_calcolato}** condomini attivi. (Quota singola: **€ {importo_totale_finale / divisore_calcolato:.2f}**)"
        )

    if st.button("💾 Salva Spesa in Archivio", type="primary"):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute(
            """INSERT INTO spese (tipo, descrizione, num_fattura, data_spesa, periodo_dal, periodo_al, importo_totale, commissione, dettagli_json, reportata, note_pagamento, num_condomini_divisore, esclusi_json)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?)""",
            (
                tipo_spesa,
                descrizione_spesa,
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
        st.success(f"✅ Spesa '{descrizione_spesa}' salvata in archivio!")
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
                    conn.commit()
                    conn.close()
                    st.warning("🗑️ Spesa eliminata!")
                    st.rerun()

    else:
        st.info("Nessuna spesa salvata in archivio.")

# ==========================================
# TAB 3: GENERAZIONE REPORT
# ==========================================
with tab3:
    st.header("3. Genera Report e Seleziona Spese da Includere")

    conn = sqlite3.connect(DB_FILE)
    df_filtrato = pd.read_sql_query("SELECT * FROM spese ORDER BY id DESC", conn)
    conn.close()

    if not df_filtrato.empty:
        st.subheader("Seleziona quali spese includere in QUESTO report:")

        df_filtrato["Includi"] = True
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
                    "Totale (€)", format="€ %.2f"
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
        )

        spese_selezionate_ids = edited_spese_rep[edited_spese_rep["Includi"] == True][
            "id"
        ].tolist()
        spese_scelte_df = df_filtrato[df_filtrato["id"].isin(spese_selezionate_ids)]

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

        st.markdown("---")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            condomino_sel = (
                st.selectbox("Seleziona Condomino per il PDF:", list_condomini)
                if list_condomini
                else None
            )
            periodo_testo = st.text_input(
                "Dicitura Intestazione per PDF:", value=intestazione_automatica
            )
        with col_c2:
            mostra_scad = st.checkbox("Inserisci Data Scadenza?", value=False)
            scad_val = (
                st.date_input(
                    "Data Scadenza", datetime.now(), format="DD/MM/YYYY"
                ).strftime("%d/%m/%Y")
                if mostra_scad
                else ""
            )

        totale_condomino = 0.0
        dettagli_pdf = []

        if condomino_sel:
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

                str_periodo = (
                    f"\nPeriodo: dal {p_dal} al {p_al}" if p_dal and p_al else ""
                )
                str_fattura = f"\nRif. {n_fatt}" if n_fatt else ""

                note_desc = row["descrizione"] + str_fattura + str_periodo
                if comm > 0:
                    note_desc += f"\n[comm. EUR {comm:.2f}]"

                if t_spesa == "Teleriscaldamento":
                    dett_perc = {}
                    try:
                        if row["dettagli_json"]:
                            dett_perc = ast.literal_eval(row["dettagli_json"])
                    except Exception:
                        dett_perc = {}
                    perc = float(dett_perc.get(condomino_sel, 0.0))
                    q = (perc / 100.0) * imp
                    dettagli_pdf.append(
                        (pulisci_testo_pdf(note_desc), f"Consumi ({perc:.1f}%)", q)
                    )
                    totale_condomino += q
                else:
                    if condomino_sel in esclusi_spesa:
                        dettagli_pdf.append(
                            (
                                pulisci_testo_pdf(note_desc),
                                "Esente da questa spesa",
                                0.0,
                            )
                        )
                    else:
                        q = imp / divisore_spesa if divisore_spesa > 0 else 0
                        dettagli_pdf.append(
                            (
                                pulisci_testo_pdf(note_desc),
                                f"Paritaria (1/{divisore_spesa})",
                                q,
                            )
                        )
                        totale_condomino += q

            st.markdown(
                f"### Importo Dovuto per {condomino_sel}: **€ {totale_condomino:.2f}**"
            )

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("👁️ Genera e Visualizza Anteprima PDF", type="primary"):

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

                    pdf.set_fill_color(44, 62, 80)
                    pdf.set_text_color(255, 255, 255)
                    pdf.set_font("Helvetica", "B", 9.5)
                    pdf.cell(
                        105, 8, " Dettaglio Servizio, Fattura e Periodo", fill=True
                    )
                    pdf.cell(45, 8, " Criterio", fill=True)
                    pdf.cell(40, 8, " Quota (EUR)", fill=True, ln=True, align="R")

                    pdf.set_text_color(50, 50, 50)
                    pdf.set_font("Helvetica", "", 9)

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
                        f"Ricevuta_{condomino_sel.replace(' ', '_')}.pdf"
                    )

            with col_btn2:
                if st.button("🔒 Segna spese come 'GIÀ REPORTATE'"):
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
                st.subheader("📄 Anteprima del Documento PDF")
                mostra_pdf_in_anteprima(st.session_state["pdf_pronto"])
                st.download_button(
                    label="💾 Scarica e Salva PDF sul Computer",
                    data=st.session_state["pdf_pronto"],
                    file_name=st.session_state["pdf_nome_file"],
                    mime="application/pdf",
                    type="primary",
                )
    else:
        st.info("Nessuna spesa trovata in archivio.")

# ==========================================
# TAB 4: GESTIONE INCASSI CONDOMINI
# ==========================================
with tab4:
    st.header("💳 Registro Incassi e Stato Pagamenti Condomini")

    col_i1, col_i2 = st.columns([1, 2])
    with col_i1:
        mesi_nomi = [
            "Gennaio",
            "Febbraio",
            "Marzo",
            "Aprile",
            "Maggio",
            "Giugno",
            "Luglio",
            "Agosto",
            "Settembre",
            "Ottobre",
            "Novembre",
            "Dicembre",
        ]
        mese_sel = st.selectbox(
            "Seleziona Mese:", mesi_nomi, index=datetime.now().month - 1
        )
        anno_sel = st.number_input(
            "Seleziona Anno:", min_value=2020, max_value=2035, value=datetime.now().year
        )
        periodo_chiave = f"{mese_sel} {anno_sel}"

    conn = sqlite3.connect(DB_FILE)
    df_spese_all = pd.read_sql_query("SELECT * FROM spese", conn)
    df_incassi_db = pd.read_sql_query(
        "SELECT * FROM incassi_condomini WHERE periodo_mese_anno = ?",
        conn,
        params=(periodo_chiave,),
    )
    conn.close()

    map_quote = {c: 0.0 for c in list_condomini}

    for _, sp in df_spese_all.iterrows():
        d_sp = parse_data_ita(sp["data_spesa"])
        if mesi_nomi[d_sp.month - 1] == mese_sel and d_sp.year == anno_sel:
            imp = float(sp["importo_totale"])
            divisore_sp = int(sp.get("num_condomini_divisore", 6) or 6)

            esclusi_sp = []
            try:
                if sp.get("esclusi_json"):
                    esclusi_sp = ast.literal_eval(sp["esclusi_json"])
            except Exception:
                esclusi_sp = []

            if sp["tipo"] == "Teleriscaldamento":
                dett_p = {}
                try:
                    if sp["dettagli_json"]:
                        dett_p = ast.literal_eval(sp["dettagli_json"])
                except Exception:
                    pass
                for c_nome in list_condomini:
                    p = float(dett_p.get(c_nome, 0.0))
                    map_quote[c_nome] += (p / 100.0) * imp
            else:
                for c_nome in list_condomini:
                    if c_nome not in esclusi_sp:
                        map_quote[c_nome] += imp / divisore_sp if divisore_sp > 0 else 0

    dati_tabella_incassi = []
    for c_nome in list_condomini:
        row_db = (
            df_incassi_db[df_incassi_db["condomino"] == c_nome]
            if not df_incassi_db.empty
            else pd.DataFrame()
        )

        is_pagato = False
        is_esente = False
        note_i = ""
        quota_calc = map_quote[c_nome]

        if not row_db.empty:
            is_pagato = bool(row_db.iloc[0]["pagato"])
            is_esente = bool(row_db.iloc[0].get("esente", 0))
            note_i = str(row_db.iloc[0]["note_incasso"] or "")

        dati_tabella_incassi.append(
            {
                "Condomino": c_nome,
                "Esente / Alloggio Vuoto": is_esente or (quota_calc == 0.0),
                "Quota Dovuta (€)": (
                    0.00 if (is_esente or quota_calc == 0.0) else round(quota_calc, 2)
                ),
                "Saldato / Pagato": is_pagato,
                "Rif. Bonifico / Note Incasso": note_i,
            }
        )

    df_incassi_editor = pd.DataFrame(dati_tabella_incassi)

    st.subheader(f"📋 Prospetto Incassi - **{periodo_chiave}**")

    edited_incassi = st.data_editor(
        df_incassi_editor,
        use_container_width=True,
        column_config={
            "Esente / Alloggio Vuoto": st.column_config.CheckboxColumn("Esente?"),
            "Quota Dovuta (€)": st.column_config.NumberColumn(
                "Quota Dovuta (€)", format="€ %.2f"
            ),
            "Saldato / Pagato": st.column_config.CheckboxColumn("Saldato?"),
            "Rif. Bonifico / Note Incasso": st.column_config.TextColumn(
                "Rif. Bonifico / Data Versamento"
            ),
        },
        disabled=["Condomino"],
        hide_index=True,
        key=f"editor_incassi_{periodo_chiave}",
    )

    edited_incassi["Quota Calcolata Finale"] = edited_incassi.apply(
        lambda r: 0.00 if r["Esente / Alloggio Vuoto"] else r["Quota Dovuta (€)"],
        axis=1,
    )

    totale_dovuto_mese = edited_incassi["Quota Calcolata Finale"].sum()
    totale_incassato_mese = edited_incassi[edited_incassi["Saldato / Pagato"] == True][
        "Quota Calcolata Finale"
    ].sum()
    totale_da_incassare = totale_dovuto_mese - totale_incassato_mese

    st.markdown("---")
    col_tot1, col_tot2, col_tot3 = st.columns(3)
    col_tot1.metric("Totale Spese Mese", f"€ {totale_dovuto_mese:.2f}")
    col_tot2.metric("Totale Incassato", f"€ {totale_incassato_mese:.2f}")
    col_tot3.metric(
        "Rimanente da Incassare",
        f"€ {totale_da_incassare:.2f}",
        delta=-totale_da_incassare if totale_da_incassare > 0 else 0,
    )

    if st.button(
        "💾 Salva Stato Incassi del Mese", type="primary", key="btn_salva_incassi"
    ):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        for _, row in edited_incassi.iterrows():
            quota_messa = (
                0.0
                if row["Esente / Alloggio Vuoto"]
                else float(row["Quota Dovuta (€)"])
            )
            c.execute(
                """INSERT INTO incassi_condomini (condomino, periodo_mese_anno, quota_dovuta, pagato, esente, note_incasso)
                         VALUES (?, ?, ?, ?, ?, ?)
                         ON CONFLICT(condomino, periodo_mese_anno) DO UPDATE SET
                            quota_dovuta=excluded.quota_dovuta,
                            pagato=excluded.pagato,
                            esente=excluded.esente,
                            note_incasso=excluded.note_incasso""",
                (
                    row["Condomino"],
                    periodo_chiave,
                    quota_messa,
                    1 if row["Saldato / Pagato"] else 0,
                    1 if row["Esente / Alloggio Vuoto"] else 0,
                    row["Rif. Bonifico / Note Incasso"],
                ),
            )
        conn.commit()
        conn.close()
        st.success(f"✅ Registrazione incassi per **{periodo_chiave}** salvata!")
        st.rerun()

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
