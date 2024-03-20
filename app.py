import io
import pandas as pd
import streamlit as st

def trova_indice_intestazione(df):
    for index, row in df.iterrows():
        for value in row:
            if isinstance(value, str) and "Style Image" in value:
                return index
    raise ValueError("Intestazione non trovata.")

def estrai_dati_excel(xls, sheet_name):
    df = xls.parse(sheet_name=sheet_name, header=None)
    header_index = trova_indice_intestazione(df)
    df = xls.parse(sheet_name=sheet_name, header=header_index)
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
    
    return df.assign(Sheet=sheet_name)

def estrai_e_riordina_dati_da_tutti_sheet(file):
    xls = pd.ExcelFile(file)
    
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
    all_extracted_data_frames = []
    for sheet_name in xls.sheet_names:
        if "cancelled" in sheet_name.lower():
            continue
        df = estrai_dati_excel(xls, sheet_name)
        colonne_taglie.update(set(df.columns) - set(colonne_fisse_prima) - set(colonne_fisse_dopo))
        all_extracted_data_frames.append(df)
    
    all_extracted_data = pd.concat(all_extracted_data_frames, ignore_index=True)
    
    ordine_completo_colonne = colonne_fisse_prima + sorted(list(colonne_taglie)) + colonne_fisse_dopo + ['Sheet']
    all_extracted_data = all_extracted_data.reindex(columns=ordine_completo_colonne)
    
    # Rimozione colonne taglie vuote
    first_size_col_idx = all_extracted_data.columns.get_loc("Country of Origin") + 1
    last_size_col_idx = all_extracted_data.columns.get_loc("Sugg. Retail (EUR)") - 1
    colonne_da_rimuovere = [col for col in all_extracted_data.columns[first_size_col_idx:last_size_col_idx + 1] if all_extracted_data[col].sum() == 0]
    all_extracted_data.drop(columns=colonne_da_rimuovere, inplace=True)
    
    all_extracted_data.drop(columns=['Style Image'], inplace=True)  # Se necessario rimuovere la colonna 'Style Image'

    return all_extracted_data

# Streamlit UI
st.title("Tabulazione JOOR")

uploaded_file = st.file_uploader("Carica il file Excel", type=['xlsx'])

if uploaded_file is not None:
    all_extracted_data = estrai_e_riordina_dati_da_tutti_sheet(uploaded_file)
    st.success("Dati estratti e riordinati con successo!")

    # Converti il DataFrame in un file Excel per il download
    towrite = io.BytesIO()
    with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
        all_extracted_data.to_excel(writer, index=False)
        writer.sheets['Sheet1'].freeze_panes(1, 0)  # Blocca la prima riga
        
    towrite.seek(0)  # Reset del puntatore

    # Crea un link per il download del file elaborato
    st.download_button(label="Scarica Excel elaborato", data=towrite, file_name="dati_elaborati.xlsx", mime="application/vnd.ms-excel
