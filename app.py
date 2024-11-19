import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter


def trova_indice_intestazione(df):
    """
    Trova la riga di intestazione cercando la colonna "Style Image".
    """
    for index, row in df.iterrows():
        for value in row:
            if isinstance(value, str) and "style image" in value.lower().strip():
                return index
    raise ValueError("Intestazione non trovata.")


def estrai_dati_excel(xls, sheet_name):
    """
    Estrae e pulisce i dati da un foglio Excel, partendo dalla riga di intestazione.
    """
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
        df[col] = df[col].replace(0, pd.NA)

    return df


def estrai_e_riordina_dati_da_tutti_sheet(xls):
    """
    Estrae e riordina i dati da tutti i fogli in un unico DataFrame.
    """
    colonne_fisse_prima = [
        "Style Image", "Style Name", "Style Number", "Color",
        "Color Code", "Color Comment", "Style Comment",
        "Materials", "Fabrication", "Country of Origin"
    ]
    colonne_fisse_dopo = [
        "Sugg. Retail (EUR)", "WholeSale (EUR)", "Item Discount",
        "Units", "Total (EUR)"
    ]
    ordine_taglie_specifiche = ["OS", "O/S", "ONE SIZE", "UNI", "XXXS", "XXS", "XS", "S", "SM", "M", "ML", "L", "XL", "XXL", "XXXL"]

    colonne_taglie = set()
    all_extracted_data_frames = []

    for sheet_name in xls.sheet_names:
        if "cancelled" in sheet_name.lower():
            continue
        try:
            df = estrai_dati_excel(xls, sheet_name)
            colonne_taglie.update(set(df.columns) - set(colonne_fisse_prima) - set(colonne_fisse_dopo))
            all_extracted_data_frames.append(df.assign(Sheet=sheet_name))
        except ValueError:
            st.warning(f"Intestazione non trovata nel foglio: {sheet_name}. Saltato.")

    all_extracted_data = pd.concat(all_extracted_data_frames, ignore_index=True)

    taglie_specifiche = [col for col in ordine_taglie_specifiche if col in colonne_taglie]
    taglie_numeriche = [col for col in colonne_taglie if col.replace('.', '', 1).isdigit() and col not in ordine_taglie_specifiche]
    altre_taglie = sorted(list(set(colonne_taglie) - set(taglie_specifiche) - set(taglie_numeriche)), key=lambda x: (x.isdigit(), x))
    taglie_numeriche_ordinate = sorted(taglie_numeriche, key=lambda x: float(x))

    ordine_completo_colonne = colonne_fisse_prima + taglie_specifiche + taglie_numeriche_ordinate + altre_taglie + colonne_fisse_dopo + ['Sheet']
    all_extracted_data = all_extracted_data.reindex(columns=ordine_completo_colonne)
    all_extracted_data = all_extracted_data[all_extracted_data["Style Image"].isna()]
    all_extracted_data.drop(columns=['Style Image'], inplace=True)

    return all_extracted_data


def get_excel_column_letter(col_idx):
    """
    Converte un indice di colonna numerico in una lettera per Excel.
    """
    letter = ''
    while col_idx > 25:
        col_idx, remainder = divmod(col_idx, 26)
        letter = chr(65 + remainder) + letter
        col_idx -= 1
    letter = chr(65 + col_idx) + letter
    return letter


def main():
    st.title("JOOR > Excel")
    st.image('joor-logo-app.jpg')

    uploaded_file = st.file_uploader("Trascina qui il tuo file Excel o clicca per caricarlo", type=['xlsx'])
    if uploaded_file is not None:
        original_file_name = uploaded_file.name.rsplit('.', 1)[0]

        xls = pd.ExcelFile(uploaded_file)
        all_extracted_data = estrai_e_riordina_dati_da_tutti_sheet(xls)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            all_extracted_data.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            worksheet.freeze_panes(1, 0)
            format_yellow = workbook.add_format({'bg_color': '#FFFF99'})

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

        processed_file_name = f"{original_file_name}_processed.xlsx"
        
        st.success("Elaborazione completata!")
        st.download_button(
            label="Scarica Excel Elaborato",
            data=output.getvalue(),
            file_name=processed_file_name,
            mime="application/vnd.ms-excel"
        )


if __name__ == "__main__":
    main()
