import pandas as pd
import streamlit as st
from io import BytesIO

def clean_and_extract_product_data(input_file):
    xls = pd.ExcelFile(input_file)
    all_data_frames = []  # Lista per mantenere tutti i DataFrame puliti
    size_columns = set()  # Set per tracciare univocamente tutte le colonne delle taglie trovate
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
        df.dropna(axis=1, how='all', inplace=True)

        start_index = None
        end_index = None
        for index, row in df.iterrows():
            if start_index is None and "Style Name" in row.values:
                start_index = index
            elif "Total:" in row.values:
                end_index = index
                break
        
        if start_index is not None and end_index is not None:
            # Imposta la prima riga valida come intestazione
            product_data_df = df.iloc[start_index:end_index]
            product_data_df.columns = product_data_df.iloc[0]
            product_data_df = product_data_df[1:]
            product_data_df.reset_index(drop=True, inplace=True)

            # Identifica colonne delle taglie
            if "Country of Origin" in product_data_df and "Sugg. Retail (EUR)" in product_data_df:
                size_start = product_data_df.columns.get_loc("Country of Origin") + 1
                size_end = product_data_df.columns.get_loc("Sugg. Retail (EUR)")
                sizes = product_data_df.columns[size_start:size_end]
                size_columns.update(sizes)
            
            all_data_frames.append(product_data_df)

    if not all_data_frames:
        return pd.DataFrame()

    # Unifica tutte le colonne per assicurare coerenza tra i DataFrame
    unified_columns = list(all_data_frames[0].columns)
    for size in sorted(size_columns):  # Assicura che le taglie siano in ordine
        if size not in unified_columns:
            unified_columns.insert(unified_columns.index("Sugg. Retail (EUR)"), size)
    
    # Concatena tutti i DataFrame in uno, riempiendo le colonne mancanti con NaN
    final_df = pd.concat(all_data_frames).reindex(columns=unified_columns)
    
    return final_df

def save_to_excel(final_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False)
    output.seek(0)
    return output

# Interfaccia Streamlit
st.title('Unione e pulizia dati Excel in un unico foglio')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=["xlsx"])
if uploaded_file is not None:
    final_df = clean_and_extract_product_data(uploaded_file)
    
    if st.button('Genera Excel Unificato') and not final_df.empty:
        output = save_to_excel(final_df)
        st.download_button(
            label="Scarica Excel Unificato",
            data=output,
            file_name="excel_unificato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    elif final_df.empty:
        st.warning("Il file caricato non contiene dati validi per l'unificazione.")
