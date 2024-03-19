def carica_immagini_da_excel(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        images = []
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            for col in df.columns:
                for index, row in df.iterrows():
                    cell_value = row[col]
                    if isinstance(cell_value, bytes):
                        try:
                            img = Image.open(io.BytesIO(cell_value))
                            images.append(img)
                        except Exception as e:
                            st.warning(f"Impossibile caricare l'immagine dalla riga {index} del foglio di lavoro '{sheet_name}' e colonna '{col}': {e}")
                    else:
                        st.info(f"Valore nella riga {index}, colonna '{col}': {cell_value}")
        return images
    except Exception as e:
        st.error(f"Si è verificato un errore durante il caricamento del file Excel: {e}")
        return None
