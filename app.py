import streamlit as st
import pandas as pd
from io import BytesIO
import re

def clean_sizes_column(df):
    """Rimuove 'Sizes' alla fine dei valori nella colonna 'Size'."""
    df['Size'] = df['Size'].apply(lambda x: re.sub(r'Sizes$', '', str(x).strip()))
    return df

def pivot_sizes(df):
    """Trasforma le righe dei valori 'Size' in colonne e raggruppa i dati."""
    # Pulizia dei valori 'Size' e preparazione per il pivot
    df = clean_sizes_column(df)
    
    # Creazione di un nuovo DataFrame con i valori 'Qty' per ogni 'Size' trasposti in colonne
    df_pivot = df.pivot_table(index=["Season", "Color", "Style Number", "Name"], 
                              columns='Size', 
                              values='Qty', 
                              aggfunc='sum', 
                              fill_value=0).reset_index()
    
    # Rinomina le colonne per rimuovere il nome del livello superiore 'Qty'
    df_pivot.columns.name = None  # Rimuove il nome della colonna d'indice
    df_pivot.reset_index(inplace=True)
    df_pivot.columns = [str(col) for col in df_pivot.columns]  # Assicura che i nomi delle colonne siano stringhe

    return df_pivot

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

st.title('Applicazione per la trasposizione e raggruppamento dei dati Excel')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=['xlsx'])
if uploaded_file is not None:
    df = load_data(uploaded_file)
    if not df.empty:
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
