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
            product_data_df = df.iloc[start_index+1:end_index]  # +1 per escludere la riga dell'intestazione
            product_data_df.columns = df.iloc[start_index]  # Imposta la riga di intestazione corretta
            product_data_df.reset_index(drop=True, inplace=True)

            # Rileva colonne delle taglie
            if "Country of Origin" in product_data_df.columns and "Sugg. Retail (EUR)" in product_data_df.columns:
                size_start = product_data_df.columns.get_loc("Country of Origin") + 1
                size_end = product_data_df.columns.get_loc("Sugg. Retail (EUR)")
                sizes = product_data_df.columns[size_start:size_end]
                size_columns.update(sizes)

            all_data_frames.append(product_data_df)

    if not all_data_frames:
        return pd.DataFrame()

    # Preparazione del DataFrame finale con le colonne unificate
    final_columns = all_data_frames[0].columns.tolist()
    for size in sorted(size_columns):
        if size not in final_columns:
            final_columns.insert(final_columns.index("Sugg. Retail (EUR)"), size)

    # Concatena i DataFrame assicurando che tutte le colonne corrispondano
    final_df = pd.concat([df.reindex(columns=final_columns, fill_value=0) for df in all_data_frames], ignore_index=True)

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
