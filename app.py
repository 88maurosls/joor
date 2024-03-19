import pandas as pd
import streamlit as st
from io import BytesIO

def clean_and_extract_product_data(input_file):
    xls = pd.ExcelFile(input_file)
    all_data_frames = []
    size_columns = set()  # Set per tracciare univocamente tutte le colonne delle taglie trovate

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
        df.dropna(axis=1, how='all', inplace=True)

        start_index = end_index = None
        for index, row in df.iterrows():
            if start_index is None and "Style Name" in row.values:
                start_index = index
            elif "Total:" in row.values:
                end_index = index
                break

        if start_index is not None and end_index is not None:
            df = df.iloc[start_index:end_index]
            df.columns = df.iloc[0]  # Imposta le intestazioni
            df = df[1:].reset_index(drop=True)

            # Trova le colonne delle taglie, se esistono
            if 'Country of Origin' in df.columns and 'Sugg. Retail (EUR)' in df.columns:
                sizes_start = df.columns.get_loc('Country of Origin') + 1
                sizes_end = df.columns.get_loc('Sugg. Retail (EUR)')
                sizes = df.columns[sizes_start:sizes_end]
                size_columns.update(sizes)
            
            all_data_frames.append(df)

    # Preparazione delle colonne finali escludendo le taglie per ora
    final_columns = ['Style Name', 'Style Number', 'Color', 'Color Code', 'Country of Origin']  # Aggiungi altre colonne fisse qui
    size_columns_sorted = sorted(list(size_columns))  # Ordina le colonne delle taglie
    additional_columns = ['Sugg. Retail (EUR)', 'WholeSale (EUR)', 'Units', 'Total (EUR)']  # Continua con altre colonne fisse dopo le taglie

    # Concatena tutti i DataFrame considerando l'ordine delle colonne stabilito
    final_df = pd.concat(all_data_frames, ignore_index=True)

    # Riordina il DataFrame finale per riflettere l'ordine desiderato, inclusi i nomi delle taglie tra le colonne specificate
    final_df = final_df.reindex(columns=final_columns + size_columns_sorted + additional_columns)

    return final_df

def save_to_excel(final_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False)
    output.seek(0)
    return output

# Streamlit UI
st.title('Excel Data Cleaning and Merging')

uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")
if uploaded_file is not None:
    final_df = clean_and_extract_product_data(uploaded_file)
    
    if not final_df.empty and st.button('Generate Cleaned Excel'):
        output = save_to_excel(final_df)
        st.download_button(
            label="Download Cleaned Excel",
            data=output,
            file_name="cleaned_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
