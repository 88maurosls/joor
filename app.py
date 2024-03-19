import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
from PIL import Image

def trova_indice_intestazione(df):
    for index, row in df.iterrows():
        for value in row:
            if isinstance(value, str) and "Style Image" in value:
                return index
    raise ValueError("Intestazione non trovata.")

def estrai_dati_excel(xls, sheet_name):
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    header_index = trova_indice_intestazione(df)
    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_index)
    
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    
    if 'Country of Origin' in df.columns:
        total_row_index = df[df['Country of Origin'].astype(str).str.contains("Total:", na=False)].index
        if not total_row_index.empty:
            df = df.iloc[:total_row_index[0]]
    
    taglie_columns = [col for col in df.columns if col not in [
        "Style Image", "Style Name", "Style Number", "Color", 
        "Color Code", "Color Comment", "Style Comment", 
        "Materials", "Fabrication", "Country of Origin",
        "Sugg. Retail (EUR)", "WholeSale (EUR)", "Item Discount", 
        "Units", "Total (EUR)"
    ]]
    for col in taglie_columns:
        df[col] = df[col].fillna(0)
    
    return df

def estrai_e_riordina_dati_da_tutti_sheet(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    
    colonne_fisse_prima = [
        "Style Image", "Style Name", "Style Number", "Color", 
        "Color Code", "Color Comment", "Style Comment", 
        "Materials", "Fabrication", "Country of Origin"
    ]
    colonne_fisse_dopo = [
        "Sugg. Retail (EUR)", "WholeSale (EUR)", "Item Discount", 
        "Units", "Total (EUR)"
    ]
    
    colonne_taglie = set()
    
    for sheet_name in xls.sheet_names:
        df = estrai_dati_excel(xls, sheet_name)
        colonne_taglie.update(set(df.columns) - set(colonne_fisse_prima) - set(colonne_fisse_dopo))

    all_extracted_data = pd.concat([estrai_dati_excel(xls, sheet_name) for sheet_name in xls.sheet_names], ignore_index=True)
    
    ordine_completo_colonne = colonne_fisse_prima + sorted(list(colonne_taglie)) + colonne_fisse_dopo

    for col in colonne_taglie:
        all_extracted_data[col] = all_extracted_data[col].fillna(0)

    all_extracted_data = all_extracted_data.reindex(columns=ordine_completo_colonne)

    return all_extracted_data

# Funzione per caricare le immagini dal file Excel
def carica_immagini_da_excel(file_path):
    wb = load_workbook(file_path)
    sheet = wb.active  # Seleziona il foglio attivo
    image_loader = SheetImageLoader(sheet)
    images = []
    for cell in sheet.iter_rows(min_row=1, max_row=1):  # Itera sulla prima riga per trovare le celle con immagini
        for col in cell:
            if image_loader.image_in(col.coordinate):
                img = image_loader.get(col.coordinate)
                images.append(img)
    return images

# Funzione per visualizzare le immagini in Streamlit
def visualizza_immagini(images):
    if images:
        st.header("Anteprima Immagini")
        for img in images:
            st.image(img, caption='Anteprima immagine')

# Streamlit UI
st.title("Tabulazione JOOR")

uploaded_file = st.file_uploader("Carica il file Excel", type=['xlsx'])

if uploaded_file is not None:
    all_extracted_data = estrai_e_riordina_dati_da_tutti_sheet(uploaded_file)
    st.success("Dati estratti e riordinati con successo!")

    # Carica le immagini dal file Excel
    images = carica_immagini_da_excel(uploaded_file)

    # Visualizza le immagini in Streamlit
    visualizza_immagini(images)

    # Converti il DataFrame in un file Excel per il download
    towrite = io.BytesIO()
    with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
        all_extracted_data.to_excel(writer, index=False)
    towrite.seek(0)  # Reset del puntatore

    # Crea un link per il download del file elaborato
    st.download_button(label="Scarica Excel elaborato", data=towrite, file_name="dati_elaborati.xlsx", mime="application/vnd.ms-excel")
