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

def is_numeric_column(col):
    # Controlla se la colonna è numericamente ordinabile, escludendo colonne non destinate all'ordinamento
    try:
        float(col)
        return True
    except ValueError:
        return False

def extract_numeric_part(col):
    # Assicurati che col sia una stringa
    col_str = str(col)
    try:
        # Estrai solo i numeri (e il punto per i decimali) dall'etichetta della colonna
        numeric_part = ''.join(filter(str.isdigit, col_str)) or '0'
        return int(numeric_part)
    except ValueError:
        # In caso di qualsiasi errore, restituisci un numero molto alto per posizionare la colonna alla fine
        return float('inf')

def save_combined_data_to_excel(cleaned_data):
    combined_df = pd.DataFrame()

    for sheet_name, data_df in cleaned_data.items():
        data_df['Sheet'] = sheet_name
        combined_df = pd.concat([combined_df, data_df], ignore_index=True)

    # Ottiene gli indici delle colonne chiave
    try:
        index_of_country_of_origin = combined_df.columns.get_loc("Country of Origin")
        index_of_sugg_retail = combined_df.columns.get_loc("Sugg. Retail (EUR)")
    except KeyError as e:
        raise KeyError(f"Colonna non trovata: {e}")

    # Divide le colonne in tre gruppi: prima, durante (taglie), e dopo
    fixed_columns_before = combined_df.columns[:index_of_country_of_origin + 1].tolist()
    size_columns = combined_df.columns[index_of_country_of_origin + 1:index_of_sugg_retail].tolist()
    fixed_columns_after = combined_df.columns[index_of_sugg_retail:].tolist()

    # Stampa delle taglie
    print("Taglie finali:", size_columns)

    # Preparazione dell'ordinamento delle colonne delle taglie
    size_columns.sort(key=lambda col: extract_numeric_part(str(col)))

    # Riorganizza il DataFrame con l'ordine desiderato delle colonne
    ordered_columns = fixed_columns_before + size_columns + fixed_columns_after
    combined_df = combined_df[ordered_columns]

    # Salvataggio in un nuovo file Excel
    output_combined = BytesIO()
    with pd.ExcelWriter(output_combined, engine='openpyxl') as writer:
        combined_df.to_excel(writer, index=False)
    output_combined.seek(0)
    return output_combined


# Interfaccia Streamlit
st.title('Unione e salvataggio dati prodotto da Excel')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=["xlsx"])
if uploaded_file is not None:
    cleaned_data = clean_and_extract_product_data(uploaded_file)
    
    if st.button('Unisci e Salva Dati'):
        output_combined = save_combined_data_to_excel(cleaned_data)
        st.download_button(
            label="Scarica Dati Uniti come Excel",
            data=output_combined,
            file_name="dati_uniti.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
