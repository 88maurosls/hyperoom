import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Funzione per pulire la colonna "Size"
def clean_sizes_column(df, size_col='Size'):
    df[size_col] = df[size_col].apply(lambda x: re.sub(r'Sizes$', '', str(x).strip()))
    return df

# Funzione per rimuovere il trattino finale dalla colonna "Style Number"
def clean_style_number(df):
    df['Style Number'] = df['Style Number'].astype(str).str.rstrip('-')
    return df

# Funzione per trasformare le righe dei valori "Size" in colonne e raggruppare i dati
def pivot_sizes(df):
    df = clean_sizes_column(df)
    df = clean_style_number(df)
    
    # Rimuovi la colonna 'Image' se presente
    if 'Image' in df.columns:
        df.drop('Image', axis=1, inplace=True)

    # Pivot della colonna 'Size'
    df_pivot = df.pivot_table(index=["Season", "Color", "Style Number", "Name"], 
                              columns='Size', 
                              values='Qty', 
                              aggfunc='sum').reset_index()

    # Rimuovi le colonne con tutti i valori NaN o 0
    df_pivot.dropna(axis=1, how='all', inplace=True)
    df_pivot.replace({0: None}, inplace=True)
    
    # Ordinamento delle colonne delle taglie
    predefined_size_order = ["OS", "O/S", "One size", "UNI", "XXXS", "XXS", "XS", "XS/S", "S", "S/M", "M", 
                             "M/L", "L", "L/XL", "XL", "XXL", "XXXL"]
    
    # Ottieni tutte le colonne delle taglie dal pivot, escludendo le altre colonne
    all_size_cols = df_pivot.columns.difference(df.columns.difference(['Size', 'Qty'])).tolist()
    size_columns = sorted(
        [col for col in all_size_cols if col in predefined_size_order],
        key=lambda x: predefined_size_order.index(x)
    ) + sorted(
        [col for col in all_size_cols if col not in predefined_size_order and not col.isdigit()],
        key=str
    ) + sorted(
        [col for col in all_size_cols if col.isdigit()],
        key=int
    )
    
    # Aggiungi le colonne di dimensioni ordinate al DataFrame finale
    df_final = df.join(df_pivot[size_columns], on=["Season", "Color", "Style Number", "Name"])
    
    # Assicurati che le date siano in formato datetime
    if 'Ship Start' in df.columns:
        df_final['Ship Start'] = pd.to_datetime(df_final['Ship Start'])
    if 'Ship End' in df.columns:
        df_final['Ship End'] = pd.to_datetime(df_final['Ship End'])

    return df_final

# Funzione per convertire il DataFrame in un oggetto Excel
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', datetime_format='yyyy-mm-dd') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    return output.getvalue()

# Funzione per caricare i dati da un file Excel
def load_data(file_path):
    return pd.read_excel(file_path)

# Titolo dell'app Streamlit
st.title('Applicazione per la trasposizione e raggruppamento dei dati Excel')

# Caricamento del file
uploaded_file = st.file_uploader("Carica il tuo file Excel", type=['xlsx'])

# Se un file è stato caricato, elaboralo
if uploaded_file is not None:
    df = load_data(uploaded_file)
    if not df.empty:
        st.write("Anteprima dei dati originali:", df)
        df_final = pivot_sizes(df)
        st.write("Anteprima dei dati trasformati:", df_final)
        processed_data = convert_df_to_excel(df_final)
        st.download_button(
            label="📥 Scarica dati Excel trasformati",
            data=processed_data,
            file_name='dati_trasformati.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.error("Il DataFrame caricato è vuoto. Si prega di caricare un file con i dati.")
else:
    st.info("Attendere il caricamento di un file Excel.")
