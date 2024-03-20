import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter

# Le funzioni per l'estrazione dei dati...

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
        df[col] = df[col].replace(0, pd.NA) # Cambia qui per usare pd.NA invece di 0
    return df

def estrai_e_riordina_dati_da_tutti_sheet(xls):
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
        all_extracted_data_frames.append(df.assign(Sheet=sheet_name))
    all_extracted_data = pd.concat(all_extracted_data_frames, ignore_index=True)
    ordine_completo_colonne = colonne_fisse_prima + sorted(list(colonne_taglie)) + colonne_fisse_dopo + ['Sheet']
    all_extracted_data = all_extracted_data.reindex(columns=ordine_completo_colonne)
    all_extracted_data = all_extracted_data[all_extracted_data["Style Image"].isna()]
    all_extracted_data.drop(columns=['Style Image'], inplace=True)  # Questa riga rimuove la colonna 'Style Image'
    return all_extracted_data

# La funzione per convertire indice colonna in lettera colonna Excel.
def get_excel_column_letter(col_idx):
    letter = ''
    while col_idx > 25:
        col_idx, remainder = divmod(col_idx, 26)
        letter = chr(65 + remainder) + letter
        col_idx -= 1
    letter = chr(65 + col_idx) + letter
    return letter

# L'applicazione principale Streamlit.
def main():
    st.title("Elaboratore di Excel")

    uploaded_file = st.file_uploader("Trascina qui il tuo file Excel o clicca per caricarlo", type=['xlsx'])
    if uploaded_file is not None:
        # Memorizza il nome del file originale senza l'estensione
        original_file_name = uploaded_file.name.rsplit('.', 1)[0]

        xls = pd.ExcelFile(uploaded_file)
        all_extracted_data = estrai_e_riordina_dati_da_tutti_sheet(xls)
        
        # Verifica se la colonna 'Style Image' esiste prima di tentare di rimuoverla
        if 'Style Image' in all_extracted_data.columns:
            all_extracted_data.drop(columns=['Style Image'], inplace=True)

        # Rimuovi le colonne delle taglie che contengono solo valori zero o NaN
        taglie_columns = [col for col in all_extracted_data.columns if col not in [
            "Style Image", "Style Name", "Style Number", "Color",
            "Color Code", "Color Comment", "Style Comment",
            "Materials", "Fabrication", "Country of Origin",
            "Sugg. Retail (EUR)", "WholeSale (EUR)", "Item Discount",
            "Units", "Total (EUR)", "Sheet"
        ]]
        for col in taglie_columns:
            if all(pd.isna(all_extracted_data[col])) or all(all_extracted_data[col] == 0):
                all_extracted_data.drop(columns=[col], inplace=True)

        # Converti DataFrame in un file Excel in memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            all_extracted_data.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            worksheet.freeze_panes(1, 0)
            format_yellow = workbook.add_format({'bg_color': '#FFFF99'})

            # Applica la formattazione condizionale solo alle colonne delle taglie
            first_size_col_idx = all_extracted_data.columns.get_loc("Country of Origin") + 1
            last_size_col_idx = all_extracted_data.columns.get_loc("Sugg. Retail (EUR)") - 1
            
            for col_idx in range(first_size_col_idx, last_size_col_idx + 1):
                col_letter = get_excel_column_letter(col_idx)
                cell_range = f'{col_letter}2:{col_letter}{len(all_extracted_data) + 1}'
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '>',
                    'value': 0,
                    'format': format_yellow
                })
                
        # Configurazione per il download del file Excel elaborato
        processed_file_name = f"{original_file_name}_processed.xlsx"  # Aggiungi "_processed" al nome del file
        
        st.success("Elaborazione completata!")
        st.download_button(
            label="Scarica Excel Elaborato",
            data=output.getvalue(),
            file_name=processed_file_name,  # Usa il nome del file modificato qui
            mime="application/vnd.ms-excel"
        )

if __name__ == "__main__":
    main()
