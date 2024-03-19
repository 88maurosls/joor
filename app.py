import pandas as pd
import streamlit as st
from io import BytesIO

def clean_and_extract_product_data(input_file):
    xls = pd.ExcelFile(input_file)
    all_data_frames = []
    size_columns = set()

    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            
            # Trova indice di "Country of Origin" e "Sugg. Retail (EUR)"
            country_of_origin_index = None
            sugg_retail_index = None
            for col_index, col in enumerate(df.columns):
                if col.strip() == "Country of Origin":
                    country_of_origin_index = col_index
                elif col.strip() == "Sugg. Retail (EUR)":
                    sugg_retail_index = col_index
                elif country_of_origin_index is not None and sugg_retail_index is not None:
                    break

            if country_of_origin_index is not None and sugg_retail_index is not None:
                # Seleziona le colonne tra "Country of Origin" e "Sugg. Retail (EUR)"
                size_cols = df.columns[country_of_origin_index + 1:sugg_retail_index]
                size_columns.update(size_cols)

                # Rimuovi le colonne delle taglie dal DataFrame
                df = df.drop(columns=size_cols)

                # Aggiungi il DataFrame pulito alla lista
                all_data_frames.append(df)
            else:
                st.warning(f"'Country of Origin' or 'Sugg. Retail (EUR)' not found in sheet: {sheet_name}")
        except Exception as e:
            st.warning(f"Error processing sheet '{sheet_name}': {str(e)}")

    if not all_data_frames:
        return pd.DataFrame()

    # Concatena tutti i DataFrame puliti
    final_df = pd.concat(all_data_frames, ignore_index=True)

    return final_df




    # Prepare final DataFrame
    final_columns = set(df for df in all_data_frames[0].columns)
    for sizes in size_columns:
        final_columns.add(sizes)
    final_columns = list(final_columns)

    # Concatenate DataFrames
    final_df = pd.concat(all_data_frames, ignore_index=True)
    final_df = final_df.reindex(columns=final_columns)

    return final_df


    # Prepare final DataFrame
    final_columns = set(df for df in all_data_frames[0].columns)
    for sizes in size_columns:
        final_columns.add(sizes)
    final_columns = list(final_columns)

    # Concatenate DataFrames
    final_df = pd.concat(all_data_frames, ignore_index=True)
    final_df = final_df.reindex(columns=final_columns)

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
