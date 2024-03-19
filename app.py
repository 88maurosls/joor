import streamlit as st
import pandas as pd
from PIL import Image
import base64

def get_image_from_base64(base64_string):
    decoded_image = base64.b64decode(base64_string)
    img = Image.open(io.BytesIO(decoded_image))
    return img

def main():
    st.title("Estrai Immagini da Excel")

    # Carica il file Excel
    file = st.file_uploader("Carica il file Excel", type=["xlsx"])

    if file is not None:
        df = pd.read_excel(file)

        # Mostra il dataframe
        st.write("Contenuto del file Excel:")
        st.write(df.to_html(escape=False), unsafe_allow_html=True)

        # Estrai e mostra le immagini
        for column in df.columns:
            if df[column].dtype == object:
                for i, image_data in enumerate(df[column]):
                    if isinstance(image_data, str) and image_data.startswith("data:image"):
                        st.subheader(f"Immagine dalla colonna '{column}' - Righe {i+1}")
                        image = get_image_from_base64(image_data.split(",")[1])
                        st.image(image)

if __name__ == "__main__":
    main()
