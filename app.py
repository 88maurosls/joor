import pandas as pd
import streamlit as st
from io import BytesIO

def clean_and_extract_product_data_corrected(input_file):
    xls = pd.ExcelFile(input_file)
    all_data_frames = []
    size_columns = set()

    # Definiamo l'ordine corretto delle colonne basandoci sull'analisi del file originale
    correct_order = ['Style Image', 'Style Name', 'Style Number', 'Color', 'Color Code', 'Color Comment',
                     'Style Comment', 'Materials', 'Fabrication', 'Country of Origin']  # Aggiungi il resto delle colonne secondo l'ordine desiderato
    
    # Utilizza le intestazioni dal file originale per stabilire le colonne delle taglie e altre colonne variabili
    size_col_start = correct_order.index('Country of Origin') + 1
    size_col_end = correct_order.index('Sugg. Retail (EUR)')
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Rileva e aggiorna l'elenco delle colonne delle taglie per ogni foglio
        current_sizes = [col for col in df.columns if col in correct_order[size_col_start:size_col_end]]
        size_columns.update(current_sizes)
        
        all_data_frames.append(df)

    # Unisci i DataFrame, mantenendo solo le colonne riconosciute e in ordine corretto
    final_df = pd.concat(all_data_frames, ignore_index=True)
    final_df = final_df.reindex(columns=[col for col in correct_order if col in final_df.columns] + sorted(list(size_columns)))

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
    # Assicurati che il nome della funzione qui corrisponda al nome della funzione definita
    final_df = clean_and_extract_product_data_corrected(uploaded_file)
    
    if not final_df.empty and st.button('Genera Excel Unificato'):
        output = save_to_excel(final_df)
        st.download_button(
            label="Scarica Excel Unificato",
            data=output,
            file_name="excel_unificato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
