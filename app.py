import streamlit as st
import pandas as pd
from io import BytesIO

def clean_sizes_column(df, column_name='Sizes'):
    # Assicurati che la colonna esista nel DataFrame
    if column_name in df.columns:
        # Rimuovi la parola "Sizes" se è alla fine dei valori della colonna
        df[column_name] = df[column_name].apply(lambda x: x[:-5].strip() if str(x).endswith('Sizes') else x.strip())
    return df

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    output.seek(0)  # Riporta il cursore all'inizio del buffer dopo la scrittura
    return output

st.title('Applicazione per l\'estrazione di dati Excel')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=['xlsx'])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    if not df.empty:
        df = clean_sizes_column(df)  # Pulisci la colonna Sizes per rimuovere "Sizes" alla fine
        st.write("Anteprima dei dati corretti:", df)
        processed_data = convert_df_to_excel(df)  # Converti il DataFrame pulito in Excel
        
        st.download_button(
            label="📥 Scarica dati Excel",
            data=processed_data,
            file_name='dati_processati.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.error("Il DataFrame caricato è vuoto. Si prega di caricare un file con i dati.")
