import pandas as pd
from io import BytesIO
import streamlit as st

def clean_and_concatenate_product_data(input_file):
    xls = pd.ExcelFile(input_file)
    all_data_frames = []
    size_columns = set()  # Per tenere traccia delle colonne delle taglie

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Identifica le colonne delle taglie
        start_col = df.columns.get_loc("Country of Origin") + 1
        end_col = df.columns.get_loc("Sugg. Retail (EUR)")
        sizes = df.columns[start_col:end_col]
        size_columns.update(sizes)
        
        all_data_frames.append(df)

    # Determina tutte le colonne uniche tra i fogli, mantenendo l'ordine
    unique_columns = list(pd.concat([df.iloc[:, :start_col], df.iloc[:, end_col:]] for df in all_data_frames).columns.drop_duplicates())
    size_columns = sorted(list(size_columns))  # Ordina le colonne delle taglie

    # Inserisce le colonne delle taglie prima di "Sugg. Retail (EUR)"
    insert_pos = unique_columns.index("Sugg. Retail (EUR)")
    final_columns = unique_columns[:insert_pos] + size_columns + unique_columns[insert_pos:]
    
    # Uniforma le intestazioni e concatena tutti i DataFrame
    final_df = pd.concat([df.reindex(columns=final_columns) for df in all_data_frames], ignore_index=True)
    
    return final_df

def save_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

# Interfaccia Streamlit
st.title('Unione dati prodotti in un unico foglio Excel')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=["xlsx"])
if uploaded_file is not None:
    final_df = clean_and_concatenate_product_data(uploaded_file)
    
    if st.button('Genera Excel Unificato'):
        output = save_to_excel(final_df)
        st.download_button(
            label="Scarica Excel Unificato",
            data=output,
            file_name="excel_unificato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
