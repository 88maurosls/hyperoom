import streamlit as st
import pandas as pd
from io import BytesIO
import re

def clean_sizes_column(df):
    """Rimuove 'Sizes' alla fine dei valori nella colonna 'Size'."""
    df['Size'] = df['Size'].apply(lambda x: re.sub(r'Sizes$', '', str(x).strip()))
    return df

def pivot_sizes(df, index_cols, values_col='Qty'):
    """Trasforma le righe dei valori 'Size' in colonne e raggruppa i dati."""
    # Pulizia dei valori 'Size' e preparazione per il pivot
    df = clean_sizes_column(df)

    # Trova tutte le colonne che non sono 'Size' o 'Qty' e che non fanno parte delle colonne d'indice
    other_cols = df.columns.difference(index_cols + ['Size', values_col]).tolist()
    
    # Creazione del pivot table
    df_pivot = df.pivot_table(index=index_cols, 
                              columns='Size', 
                              values=values_col, 
                              aggfunc='sum', 
                              fill_value=0).reset_index()

    # Concatenazione con le altre colonne
    df_pivot = df_pivot.join(df.set_index(index_cols)[other_cols].drop_duplicates())

    # Riassembla il multiindex nel nome delle colonne per le colonne pivotate
    df_pivot.columns = ['_'.join(col).rstrip('_') if isinstance(col, tuple) else col for col in df_pivot.columns.values]
    
    return df_pivot

def convert_df_to_excel(df):
    """Converti il DataFrame in un oggetto Excel e restituisci il buffer."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    return output.getvalue()

def load_data(uploaded_file):
    """Carica i dati da un file Excel caricato."""
    return pd.read_excel(uploaded_file)

st.title('Applicazione per la trasposizione e raggruppamento dei dati Excel')

uploaded_file = st.file_uploader("Carica il tuo file Excel", type=['xlsx'])
if uploaded_file is not None:
    df = load_data(uploaded_file)
    if not df.empty:
        # Aggiungi qui tutte le colonne che intendi usare come indice
        index_columns = ['Season', 'Color', 'Style Number', 'Name']
        df_final = pivot_sizes(df, index_columns)
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
