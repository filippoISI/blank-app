############################################################
# app.py - Esempio Streamlit con logica Pianificazione
############################################################

import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime

# ------------------------------------------
# 1) AUTENTICAZIONE GOOGLE SHEETS
# ------------------------------------------
# Se preferisci mettere il file JSON nel repo,
# crea un file "Programmazione.json" e punta qui sotto
# Altrimenti userai st.secrets (vedi passo 4)
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ESEMPIO: caricamento credenziali da file JSON nel repo
# ATTENZIONE ai file segreti in un repo pubblico
credenziali = ServiceAccountCredentials.from_json_keyfile_name(
    "Programmazione.json",  # <-- Sostituisci col tuo file, se lo metti nel repo
    scope
)
client = gspread.authorize(credenziali)

# ID del foglio Google (prendi tra /d/ e /edit nell'URL)
doc_id = "13c4FPvPL5X5DaphkGlPlY7Oxi51FepGyQIOq4OWvNFo"  # <-- Cambia col tuo
foglio_google = client.open_by_key(doc_id)

# Recupero tutti i "tab" (worksheet) interni
all_worksheets = foglio_google.worksheets()
sheet_names = [ws.title for ws in all_worksheets]
diz_sheets = {ws.title: ws for ws in all_worksheets}

# ------------------------------------------
# 2) FUNZIONI LOGICA (collisioni, skip weekend, ecc.)
# ------------------------------------------

def skip_weekend(dt: datetime.datetime) -> datetime.datetime:
    while dt.weekday() >= 5:  # sab=5, dom=6
        dt += datetime.timedelta(days=1)
    return dt

def converti_valore_in_ora(val):
    if isinstance(val, datetime.datetime):
        return val.hour
    if isinstance(val, (int, float)):
        return int(val)
    if isinstance(val, str):
        for fmt in ("%H:%M", "%H.%M", "%H:%M:%S", "%H.%M:%S"):
            try:
                return datetime.datetime.strptime(val.strip(), fmt).hour
            except:
                continue
    return None

def get_cell_value(ws, r, c):
    return ws.cell(r, c).value

def set_cell_value(ws, r, c, value):
    ws.update_cell(r, c, value)

def trova_riga_data(ws, data_cercata: datetime.date):
    colonna_date = ws.col_values(1)
    for idx, cell_val in enumerate(colonna_date, start=1):
        if not cell_val.strip():
            continue
        for fmt in ('%d/%m/%y', '%d/%m/%Y', '%Y-%m-%d'):
            try:
                dt_parsed = datetime.datetime.strptime(cell_val.strip(), fmt)
                if dt_parsed.date() == data_cercata:
                    return idx
                break
            except:
                pass
    return None

def trova_colonna_ora(ws, ora_cercata: int):
    riga2 = ws.row_values(2)
    for col_idx, val in enumerate(riga2, start=1):
        if col_idx == 1:  # skip colonna A
            continue
        if converti_valore_in_ora(val) == ora_cercata:
            return col_idx
    return None

def trova_ora_inizio_giornata(ws):
    riga2 = ws.row_values(2)
    ore = []
    for i, val in enumerate(riga2, start=1):
        if i == 1:
            continue
        ora = converti_valore_in_ora(val)
        if ora is not None:
            ore.append(ora)
    if ore:
        return min(ore)
    return 8

def trasla_ordine(ws, riga, colonna, inizio_giornata):
    ordine_da_spostare = get_cell_value(ws, riga, colonna)
    set_cell_value(ws, riga, colonna, "")
    data_string = get_cell_value(ws, riga, 1)
    
    # converte data
    data_cella = None
    for fmt in ('%d/%m/%y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            dt_parsed = datetime.datetime.strptime(data_string.strip(), fmt)
            data_cella = dt_parsed.date()
            break
        except:
            pass
    if not data_cella:
        return
    
    ora_cella_val = get_cell_value(ws, 2, colonna)
    ora_cella = converti_valore_in_ora(ora_cella_val)
    if ora_cella is None:
        return

    dt = datetime.datetime.combine(data_cella, datetime.time(hour=ora_cella)) + datetime.timedelta(hours=1)
    dt = skip_weekend(dt)
    if dt.hour >= 18:
        dt += datetime.timedelta(days=1)
        dt = skip_weekend(dt).replace(hour=inizio_giornata)

    nuova_riga = trova_riga_data(ws, dt.date())
    nuova_colonna = trova_colonna_ora(ws, dt.hour)
    if nuova_riga and nuova_colonna:
        if get_cell_value(ws, nuova_riga, nuova_colonna):
            trasla_ordine(ws, nuova_riga, nuova_colonna, inizio_giornata)
        set_cell_value(ws, nuova_riga, nuova_colonna, ordine_da_spostare)

def inserisci_non_urgente_spezzato(ws, dt_inizio, durata, ordine, inizio_giornata):
    ore_inserite = 0
    dt_corrente = dt_inizio
    while ore_inserite < durata:
        dt_corrente = skip_weekend(dt_corrente)
        if dt_corrente.hour >= 18:
            dt_corrente += datetime.timedelta(days=1)
            dt_corrente = skip_weekend(dt_corrente).replace(hour=inizio_giornata)

        riga = trova_riga_data(ws, dt_corrente.date())
        colonna = trova_colonna_ora(ws, dt_corrente.hour)
        if not riga or not colonna:
            break

        if get_cell_value(ws, riga, colonna):
            dt_corrente += datetime.timedelta(hours=1)
        else:
            set_cell_value(ws, riga, colonna, ordine)
            ore_inserite += 1
            dt_corrente += datetime.timedelta(hours=1)

    st.success(f"Ordine '{ordine}' inserito (NON urgente) su {ore_inserite} slot.")


def inserisci_ordine(ws, data_inizio, ora_inizio, durata, ordine, urgente):
    inizio_giornata = trova_ora_inizio_giornata(ws)
    
    dt_inizio = datetime.datetime.combine(data_inizio, datetime.time(hour=ora_inizio))
    slot_necessari = []
    temp_dt = dt_inizio
    for _ in range(durata):
        if temp_dt.hour >= 18:
            temp_dt += datetime.timedelta(days=1)
            temp_dt = skip_weekend(temp_dt).replace(hour=inizio_giornata)
        slot_necessari.append(temp_dt)
        temp_dt += datetime.timedelta(hours=1)

    collisioni = False
    for dt in slot_necessari:
        riga = trova_riga_data(ws, dt.date())
        colonna = trova_colonna_ora(ws, dt.hour)
        if (not riga) or (not colonna):
            st.error("Data o ora non trovata nel foglio.")
            return
        if get_cell_value(ws, riga, colonna):
            collisioni = True
            break

    if not collisioni:
        if urgente:
            for dt in slot_necessari:
                riga = trova_riga_data(ws, dt.date())
                colonna = trova_colonna_ora(ws, dt.hour)
                set_cell_value(ws, riga, colonna, ordine)
            st.success(f"Ordine urgente '{ordine}' inserito (no collisioni).")
        else:
            inserisci_non_urgente_spezzato(ws, dt_inizio, durata, ordine, inizio_giornata)
    else:
        if urgente:
            # Traslazione
            for dt in slot_necessari:
                riga = trova_riga_data(ws, dt.date())
                colonna = trova_colonna_ora(ws, dt.hour)
                if get_cell_value(ws, riga, colonna):
                    trasla_ordine(ws, riga, colonna, inizio_giornata)
            for dt in slot_necessari:
                riga = trova_riga_data(ws, dt.date())
                colonna = trova_colonna_ora(ws, dt.hour)
                set_cell_value(ws, riga, colonna, ordine)
            st.success(f"Ordine urgente '{ordine}' inserito (con traslazione).")
        else:
            inserisci_non_urgente_spezzato(ws, dt_inizio, durata, ordine, inizio_giornata)

# ------------------------------------------
# 3) FRONTEND STREAMLIT
# ------------------------------------------
st.title("Pianificazione Produzione - DEMO")

with st.form("inserimento_ordini"):
    st.subheader("Inserisci Ordine")
    ordine = st.text_input("Nome Ordine:", "")
    data_sel = st.date_input("Data:", datetime.date.today())
    ora_sel = st.number_input("Ora di inizio:", min_value=0, max_value=23, value=8)
    durata_sel = st.number_input("Durata (ore):", min_value=1, max_value=24, value=1)
    urgente_sel = st.checkbox("Urgente", value=False)
    macchinario_sel = st.selectbox("Macchina (Foglio):", sheet_names)

    submitted = st.form_submit_button("Inserisci Ordine")
    if submitted:
        ws = diz_sheets[macchinario_sel]
        inserisci_ordine(ws, data_sel, ora_sel, durata_sel, ordine, urgente_sel)

st.write("---")

# Sezione per "cancellare" o "pulire" ordini
st.subheader("Cancella Tabella o Ordine")
col1, col2 = st.columns(2)

with col1:
    if st.button("Cancella Tabella Selezionata"):
        # Svuotare la tabella su cui siamo
        ws = diz_sheets[st.selectbox("Seleziona Macchina (per pulire)", sheet_names, key="cancella_tabella")]
        # Puliamo un range ampio
        start_row = 3
        end_row = 200
        start_col_letter = "B"
        end_col_letter = "Z"
        range_clear = f"{start_col_letter}{start_row}:{end_col_letter}{end_row}"
        empty_matrix = [[""] * 26 for _ in range(end_row - start_row + 1)]
        ws.update(range_clear, empty_matrix)
        st.success("Tabella svuotata con successo.")

with col2:
    ordine_da_cancellare = st.text_input("Nome ordine da cancellare", key="nome_ordine_da_canc")
    seleziona_foglio_canc = st.selectbox("Macchina (Foglio) per cancellazione ordine:", sheet_names, key="cancella_ordine")
    if st.button("Cancella Ordine"):
        ws = diz_sheets[seleziona_foglio_canc]
        all_data = ws.get_all_values()
        n_rows = len(all_data)
        n_cols = len(all_data[0]) if all_data else 0
        cancellato = False
        for r in range(3, n_rows+1):
            row_vals = ws.row_values(r)
            for c in range(2, n_cols+1):
                idx_in_list = c - 1
                if idx_in_list < len(row_vals) and row_vals[idx_in_list] == ordine_da_cancellare:
                    set_cell_value(ws, r, c, "")
                    cancellato = True
        if cancellato:
            st.success(f"Ordine '{ordine_da_cancellare}' cancellato con successo.")
        else:
            st.warning(f"Ordine '{ordine_da_cancellare}' non trovato.")

# Optional: visualizza i dati del foglio selezionato
st.write("---")
st.subheader("Visualizza dati di un foglio")
foglio_da_vedere = st.selectbox("Scegli Foglio da visualizzare", sheet_names, key="visualizza_foglio")
if st.button("Aggiorna Visualizzazione"):
    ws_view = diz_sheets[foglio_da_vedere]
    dati = ws_view.get_all_values()
    st.write(dati)  # Visualizzazione grezza (lista di liste)
    # Oppure potresti formattarli in un DataFrame:
    # import pandas as pd
    # df = pd.DataFrame(dati)
    # st.dataframe(df)
