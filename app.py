import io
import pandas as pd
import streamlit as st
from PIL import Image

def carica_immagini_da_excel(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        images = []
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            for col in df.columns:
                for index, row in df.iterrows():
                    if isinstance(row[col], bytes):
                        try:
                            img = Image.open(io.BytesIO(row[col]))
                            images.append(img)
                        except Exception as e:
                            st.warning(f"Impossibile caricare l'immagine dalla riga {index} del foglio di lavoro '{sheet_name}' e colonna '{col}': {e}")
        return images
    except Exception as e:
        st.error(f"Si è verificato un errore durante il caricamento del file Excel: {e}")
        return None

# Streamlit UI
st.title("Tabulazione JOOR")

uploaded_file = st.file_uploader("Carica il file Excel", type=['xlsx'])

if uploaded_file is not None:
    # Carica le immagini dal file Excel
    images = carica_immagini_da_excel(uploaded_file)

    if images:
        st.header("Anteprima Immagini")
        for img in images:
            st.image(img, caption='Anteprima immagine')
    else:
        st.warning("Nessuna immagine trovata nel file Excel.")
