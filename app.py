import streamlit as st
import pandas as pd
from io import BytesIO

def convert_df_to_excel(df):
    output = BytesIO()
    # Utilizzo il contesto di 'with' per assicurarmi che tutte le operazioni sul file siano chiuse correttamente
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        # Non è necessario chiamare writer.save() qui, perché il contesto 'with' gestisce la chiusura
    output.seek(0)  # Riportiamo il cursore del file al suo inizio dopo aver finito di scrivere
    return output

st.title('Applicazione per l\'estrazione di dati Excel')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=['xlsx'])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    if not df.empty:
        st.write("Anteprima dei dati caricati:", df)
        processed_data = convert_df_to_excel(df)  # Converti il DataFrame in Excel
        
        st.download_button(
            label="📥 Scarica dati Excel",
            data=processed_data,
            file_name='dati_processati.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.error("Il DataFrame caricato è vuoto. Si prega di caricare un file con i dati.")
