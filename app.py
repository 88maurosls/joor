import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter

def trova_indice_intestazione(df):
    for index, row in df.iterrows():
        for value in row:
            if isinstance(value, str) and "Style Image" in value:
                return index
    raise ValueError("Intestazione non trovata.")

def estrai_dati_excel(xls, sheet_name):
    df = xls.parse(sheet_name, header=None)
    header_index = trova_indice_intestazione(df)
    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_index)
    
    # Rimuovi eventuali colonne "Unnamed" subito dopo aver identificato l'intestazione
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    
    # Trova e rimuovi righe dopo "Total:" se presente
    if 'Country of Origin' in df.columns:
        total_row_index = df[df['Country of Origin'].astype(str).str.contains("Total:", na=False)].index
        if not total_row_index.empty:
            df = df.iloc[:total_row_index[0]]
    
    # Gestisci i valori mancanti nelle colonne delle taglie
    taglie_columns = [col for col in df.columns if col not in [
        "Style Image", "Style Name", "Style Number", "Color", 
        "Color Code", "Color Comment", "Style Comment", 
        "Materials", "Fabrication", "Country of Origin",
        "Sugg. Retail (EUR)", "WholeSale (EUR)", "Item Discount", 
        "Units", "Total (EUR)"
    ]]
    for col in taglie_columns:
        df[col] = df[col].fillna(0)  # Sostituisci i valori mancanti con zeri
    
    return df


def estrai_e_riordina_dati_da_tutti_sheet(file_path):
    xls = pd.ExcelFile(file_path)
    
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
        # Ignora gli sheet che contengono nel nome "cancelled"
        if "cancelled" in sheet_name.lower():
            continue
        df = estrai_dati_excel(xls, sheet_name)
        colonne_taglie.update(set(df.columns) - set(colonne_fisse_prima) - set(colonne_fisse_dopo))
        all_extracted_data_frames.append(df.assign(Sheet=sheet_name))
    
    all_extracted_data = pd.concat(all_extracted_data_frames, ignore_index=True)
    
    ordine_completo_colonne = colonne_fisse_prima + sorted(list(colonne_taglie)) + colonne_fisse_dopo + ['Sheet']

    for col in colonne_taglie:
        all_extracted_data[col] = all_extracted_data[col].fillna(0)

    all_extracted_data = all_extracted_data[all_extracted_data["Style Image"].isna()]

    all_extracted_data = all_extracted_data.reindex(columns=ordine_completo_colonne)

    return all_extracted_data

def get_excel_column_letter(col_idx):
    letter = ''
    while col_idx > 25:
        col_idx, remainder = divmod(col_idx, 26)
        letter = chr(65 + remainder) + letter
        col_idx -= 1
    letter = chr(65 + col_idx) + letter
    return letter

def main():
    st.title("Elaboratore di Excel")
    
    uploaded_file = st.file_uploader("Trascina qui il tuo file Excel o clicca per caricarlo", type=['xlsx'])
    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            all_extracted_data = estrai_e_riordina_dati_da_tutti_sheet(xls)  # Assicurati che questa funzione gestisca un oggetto ExcelFile
            
            # Implementazione delle funzioni mancanti qui
            
            # Calcola gli indici delle colonne delle taglie
            first_size_col_idx = all_extracted_data.columns.get_loc("Country of Origin") + 1
            last_size_col_idx = all_extracted_data.columns.get_loc("Sugg. Retail (EUR)") - 1

            # Identifica e rimuovi le colonne delle taglie con solo valori zero
            colonne_da_rimuovere = [col for col in all_extracted_data.columns[first_size_col_idx:last_size_col_idx + 1] if all_extracted_data[col].sum() == 0]
            all_extracted_data.drop(columns=colonne_da_rimuovere, inplace=True)
            
            all_extracted_data.drop(columns=['Style Image'], inplace=True)  # Se desideri rimuovere la colonna 'Style Image'
            
            # Converti DataFrame in un file Excel in memoria
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                all_extracted_data.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                
                # Formattazione Excel aggiuntiva qui
                worksheet.freeze_panes(1, 0)
                format_yellow = workbook.add_format({'bg_color': '#FFFF99'})
                
                for col_idx in range(first_size_col_idx, last_size_col_idx + 1):
                    col_letter = get_excel_column_letter(col_idx - 1)  # Ajuste per la conversione zero-based a uno-based
                    cell_range = f'{col_letter}2:{col_letter}{len(all_extracted_data) + 1}'
                    worksheet.conditional_format(cell_range, {
                        'type': 'cell',
                        'criteria': '!=',
                        'value': 0,
                        'format': format_yellow
                    })
            
            st.success("Elaborazione completata!")
            st.download_button(label="Scarica Excel Elaborato", 
                               data=output.getvalue(), 
                               file_name="excel_elaborato.xlsx", 
                               mime="application/vnd.ms-excel")
        except Exception as e:
            st.error(f"Errore durante l'elaborazione: {e}")

if __name__ == "__main__":
    main()
