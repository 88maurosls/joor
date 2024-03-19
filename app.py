import io
import pandas as pd
import streamlit as st
from PIL import Image

def carica_immagini_da_excel(file_path, sheet_name, colonna_immagini):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        images = []
        for index, row in df.iterrows():
            if colonna_immagini in row and isinstance(row[colonna_immagini], bytes):
                try:
                    img = Image.open(io.BytesIO(row[colonna_immagini]))
                    images.append(img)
                except Exception as e:
                    st.warning(f"Impossibile caricare l'immagine dalla riga {index}: {e}")
        return images
    except Exception as e:
        st.error(f"Si è verificato un errore durante il caricamento del file Excel: {e}")
        return None

# Streamlit UI
st.title("Tabulazione JOOR")

uploaded_file = st.file_uploader("Carica il file Excel", type=['xlsx'])

if uploaded_file is not None:
    # Imposta il nome del foglio di lavoro e la colonna in cui cercare le immagini
    sheet_name = "PO# 15289031"  # Inserisci il nome del foglio di lavoro
    colonna_immagini = "A"  # Inserisci il nome della colonna delle immagini
    
    # Carica le immagini dal file Excel
    images = carica_immagini_da_excel(uploaded_file, sheet_name, colonna_immagini)

    if images:
        st.header("Anteprima Immagini")
        for img in images:
            st.image(img, caption='Anteprima immagine')
    else:
        st.warning("Nessuna immagine trovata nel file Excel.")
