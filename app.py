import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta, timezone
import zipfile

ORARI_PREDEFINITI = {
    "MATTINO": ("08:00", "14:30"),
    "POMERIGGIO": ("14:30", "21:00"),
    "NOTTE": ("21:00", "08:00+1"),
    "OBI": ("14:30", "20:00"),
    "OB": ("08:00", "14:30"),
    "M3": ("08:15", "15:30"),
    "PONTE": ("11:30", "18:30"),
}
ORARIO_DEFAULT = ("08:00", "14:00")
INDIRIZZO = "PS San Paolo, Via San Vigilio 22 Milano Italia"

def estrai_turni(df, nome):
    turni = []
    for _, row in df.iterrows():
        data = row["Data"]
        for col in df.columns[1:]:
            val = row[col]
            if isinstance(val, str) and nome.upper() in val.strip().upper():
                tipo_turno = next((k for k in ORARI_PREDEFINITI if k in str(col).upper()), None)
                if tipo_turno:
                    start_str, end_str = ORARI_PREDEFINITI[tipo_turno]
                    titolo = f"Turno {tipo_turno.title()}"
                else:
                    start_str, end_str = ORARIO_DEFAULT
                    titolo = "Turno Generico"

                start_time = datetime.strptime(f"{data.date()} {start_str}", "%Y-%m-%d %H:%M")
                if "+1" in end_str:
                    end_time = datetime.strptime(f"{data.date() + timedelta(days=1)} {end_str.replace('+1', '')}", "%Y-%m-%d %H:%M")
                else:
                    end_time = datetime.strptime(f"{data.date()} {end_str}", "%Y-%m-%d %H:%M")

                turni.append({
                    "Titolo": titolo,
                    "Inizio": start_time,
                    "Fine": end_time
                })
    return turni

def crea_file_ics(turno, index, output_dir, nome):
    dt_fmt = "%Y%m%dT%H%M%S"
    start = turno['Inizio'].strftime(dt_fmt)
    end = turno['Fine'].strftime(dt_fmt)
    uid = f"turno{index}@{nome.lower()}"
    nome_file = f"turno_{index:02d}_{turno['Titolo'].replace(' ', '_')}.ics"

    contenuto = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//TurniMedico//EN
BEGIN:VEVENT
UID:{uid}
DTSTAMP:{datetime.now(timezone.utc).strftime(dt_fmt)}Z
DTSTART;TZID=Europe/Rome:{start}
DTEND;TZID=Europe/Rome:{end}
SUMMARY:{turno['Titolo']}
DESCRIPTION:Turno lavorativo assegnato a {nome}
LOCATION:{INDIRIZZO}
END:VEVENT
END:VCALENDAR
"""

    path_completo = os.path.join(output_dir, nome_file)
    with open(path_completo, "w") as f:
        f.write(contenuto)
    return path_completo

# --- Interfaccia Streamlit ---

st.set_page_config(page_title="Turni PS San Paolo", page_icon="favicon.png")

st.markdown(
    """
    <style>
        .stApp {
            background-image: url("background.png");
            background-size: cover;
            background-repeat: no-repeat;
            background-attachment: fixed;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Turni PS San Paolo")
st.markdown("Carica il file Excel dei turni e inserisci il nome del medico da estrarre.")

uploaded_file = st.file_uploader("File Excel", type=["xlsx"])
nome_medico = st.text_input("Nome del medico")

if uploaded_file and nome_medico:
    try:
        xls = pd.ExcelFile(uploaded_file)
        foglio_attivo = xls.sheet_names[0]
        df = pd.read_excel(uploaded_file, sheet_name=foglio_attivo, header=1)
        st.info(f"Foglio attivo: {foglio_attivo}")

        prima_colonna = df.columns[0]
        df = df[df[prima_colonna].astype(str).str.contains('2025-05', na=False)]
        df = df.rename(columns={prima_colonna: 'Data'})
        df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
        df = df[df['Data'].notna()]

        turni = estrai_turni(df, nome_medico)

        if not turni:
            st.warning("Nessun turno trovato per il medico indicato.")
        else:
            st.success(f"Trovati {len(turni)} turni per {nome_medico}.")

            with st.spinner("Generazione dei file ICS..."):
                output_dir = "ics_files"
                os.makedirs(output_dir, exist_ok=True)

                ics_paths = [crea_file_ics(t, i+1, output_dir, nome_medico) for i, t in enumerate(turni)]
                zip_path = f"{output_dir}_{nome_medico}.zip"
                with zipfile.ZipFile(zip_path, 'w') as z:
                    for file in ics_paths:
                        z.write(file, os.path.basename(file))

                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="Scarica i turni in formato ZIP",
                        data=f,
                        file_name=f"turni_{nome_medico}.zip",
                        mime="application/zip"
                    )

    except Exception as e:
        st.error(f"Errore durante l'elaborazione: {e}")
