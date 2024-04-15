import streamlit as st
import pandas as pd
from io import BytesIO

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.save()
    processed_data = output.getvalue()
    return processed_data

st.title('Applicazione per l\'estrazione di dati Excel')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=['xlsx'])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Anteprima dei dati caricati:", df)
    processed_data = convert_df_to_excel(df)  # Converti il DataFrame in Excel

    st.download_button(
        label="📥 Scarica dati Excel",
        data=processed_data,
        file_name='dati_processati.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
