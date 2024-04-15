import streamlit as st
import pandas as pd
from io import BytesIO
import re

def clean_sizes_column(df, size_col='Size'):
    """Rimuove 'Sizes' alla fine dei valori nella colonna 'Size'."""
    df[size_col] = df[size_col].apply(lambda x: re.sub(r'Sizes$', '', str(x).strip()))
    return df

def pivot_sizes(df):
    """Trasforma le righe dei valori 'Size' in colonne e raggruppa i dati."""
    # Pulizia dei valori 'Size'
    df = clean_sizes_column(df)
    
    # Identificazione delle colonne che non saranno trasformate
    non_pivot_cols = df.columns.difference(['Size', 'Qty']).tolist()

    # Creazione del DataFrame pivotato
    df_pivot = df.pivot_table(index=["Season", "Color", "Style Number", "Name"], 
                              columns='Size', 
                              values='Qty', 
                              aggfunc='sum', 
                              fill_value=0).reset_index()

    # Combina i nomi delle colonne multi-livello in uno
    df_pivot.columns = [' '.join(col).strip() if type(col) is tuple else col for col in df_pivot.columns.values]

    # Estrazione delle colonne delle taglie dal DataFrame pivotato
    size_columns = df_pivot.columns.difference(non_pivot_cols)

    # Unione del pivot con le altre colonne non pivotate mantenendo l'ordine delle colonne originale
    df_final = pd.merge(df[non_pivot_cols].drop_duplicates(), df_pivot, on=["Season", "Color", "Style Number", "Name"], how='right')

    # Ordina le colonne mettendo le colonne delle taglie alla fine
    final_columns = non_pivot_cols + list(size_columns)
    df_final = df_final[final_columns]

    return df_final



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
