import pandas as pd
import streamlit as st
from io import BytesIO

def clean_and_extract_product_data(input_file):
    xls = pd.ExcelFile(input_file)
    sheet_names = xls.sheet_names

    cleaned_data = {}
    
    for sheet_name in sheet_names:
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
            product_data_df = df.iloc[start_index:end_index]
            product_data_df.columns = product_data_df.iloc[0]
            product_data_df = product_data_df[1:]
            product_data_df.reset_index(drop=True, inplace=True)
            
            cleaned_data[sheet_name] = product_data_df
        else:
            st.warning(f"Non è stato possibile trovare i dati degli oggetti acquistati nel foglio: {sheet_name}")
    
    return cleaned_data

def save_cleaned_data_to_excel(cleaned_data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, data_df in cleaned_data.items():
            data_df.to_excel(writer, sheet_name=sheet_name)
    output.seek(0)  # Sposta il cursore all'inizio del file per il download
    return output

# Interfaccia Streamlit
st.title('Pulizia e estrazione dati prodotto da Excel')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=["xlsx"])
if uploaded_file is not None:
    cleaned_data = clean_and_extract_product_data(uploaded_file)
    
    if st.button('Genera Dati Puliti'):
        output = save_cleaned_data_to_excel(cleaned_data)
        st.download_button(
            label="Scarica Dati Puliti come Excel",
            data=output,
            file_name="data_puliti.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
