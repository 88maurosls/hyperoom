import streamlit as st
import pandas as pd
from io import BytesIO
import re

def clean_sizes_column(df, column_name='Size'):
    """Pulizia della colonna specificata rimuovendo 'Sizes' alla fine dei valori."""
    if column_name in df.columns:
        df[column_name] = df[column_name].apply(lambda x: re.sub(r'Sizes$', '', str(x).strip()))
    return df

def convert_df_to_excel(df):
    """Converti il DataFrame in un oggetto Excel e restituisci il buffer."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    return output.getvalue()

def load_data(file_path):
    """Carica i dati da un file Excel specificato."""
    return pd.read_excel(file_path)

def filter_qty(df, qty_column='Qty'):
    """Filtro le righe basate sulla colonna 'Qty' per escludere valori nulli o zero."""
    return df[df[qty_column].notna() & (df[qty_column] != 0)]

st.title('Applicazione per l\'estrazione e pulizia dei dati Excel')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=['xlsx'])
if uploaded_file is not None:
    df = load_data(uploaded_file)
    if not df.empty:
        st.write("Anteprima dei dati originali:", df)
        df_cleaned = clean_sizes_column(df)
        df_filtered = filter_qty(df_cleaned)  # Filtra le righe dove 'Qty' è null o zero
        st.write("Anteprima dei dati puliti:", df_filtered)
        processed_data = convert_df_to_excel(df_filtered)
        st.download_button(
            label="📥 Scarica dati Excel puliti",
            data=processed_data,
            file_name='dati_puliti.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.error("Il DataFrame caricato è vuoto. Si prega di caricare un file con i dati.")
else:
    st.info("Attendere il caricamento di un file Excel.")
