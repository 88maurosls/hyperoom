import streamlit as st
import pandas as pd
from io import BytesIO
import re

def clean_sizes_column(df, size_col='Size'):
    """Rimuove 'Sizes' alla fine dei valori nella colonna 'Size'."""
    df[size_col] = df[size_col].apply(lambda x: re.sub(r'Sizes$', '', str(x).strip()))
    return df

def clean_style_number(df):
    """Rimuove il trattino finale dalla colonna 'Style Number' se presente."""
    df['Style Number'] = df['Style Number'].str.rstrip('-')
    return df

def pivot_sizes(df):
    """Trasforma le righe dei valori 'Size' in colonne, ordina le colonne come specificato, e rimuove 'Image'."""
    # Pulizia dei valori 'Size' e 'Style Number'
    df = clean_sizes_column(df)
    df = clean_style_number(df)
    
    # Rimozione della colonna 'Image' se esiste
    if 'Image' in df.columns:
        df.drop('Image', axis=1, inplace=True)

    # Rimozione della colonna 'Total Price (EUR)' se esiste
    if 'Total Price (EUR)' in df.columns:
        df.drop('Total Price (EUR)', axis=1, inplace=True)

    # Rimozione della colonna 'Total Units' se esiste
    if 'Total Units' in df.columns:
        df.drop('Total Units', axis=1, inplace=True)

    # Rimozione della colonna 'Units per pack' se esiste
    if 'Units per pack' in df.columns:
        df.drop('Units per pack', axis=1, inplace=True)
    
    # Creazione del DataFrame pivotato
    df_pivot = df.pivot_table(index=["Season", "Color", "Style Number", "Name"], 
                              columns='Size', 
                              values='Qty', 
                              aggfunc='sum').reset_index()

    # Combina i nomi delle colonne multi-livello in uno
    df_pivot.columns = [' '.join(col).strip() if isinstance(col, tuple) else col for col in df_pivot.columns.values]

    # Sostituzione degli zeri con NaN (o puoi usare None per null)
    df_pivot.replace({0: None}, inplace=True)

    # Rimozione delle colonne delle taglie che contengono solo valori null
    df_pivot.dropna(axis=1, how='all', inplace=True)

    # Definizione dell'ordine delle taglie
    predefined_size_order = ["OS", "O/S", "One size", "UNI", "XXXS", "XXS", "XS", "XS/S", "S", "S/M", "M", 
                             "M/L", "L", "L/XL", "XL", "XXL", "XXXL"]
    size_columns = [col for col in df_pivot.columns if col not in df.columns]

    # Ordinamento delle colonne delle taglie
    predefined_sizes = [size for size in predefined_size_order if size in size_columns]
    undefined_sizes = [size for size in size_columns if size not in predefined_size_order]

    # Dividi ulteriormente undefined_sizes in numeriche e non numeriche e ordina
    numeric_sizes = sorted([size for size in undefined_sizes if size.isdigit()], key=int)
    non_numeric_sizes = sorted([size for size in undefined_sizes if not size.isdigit()])

    # Ordine finale delle taglie
    final_size_order = predefined_sizes + non_numeric_sizes + numeric_sizes

    # Unione del pivot con le altre colonne non pivotate
    non_pivot_cols = df.columns.difference(['Size', 'Qty']).tolist()
    df_final = pd.merge(df[non_pivot_cols].drop_duplicates(), df_pivot, 
                        on=["Season", "Color", "Style Number", "Name"], how='right')

    # Predefinire l'ordine delle colonne principali
    main_cols_order = ["Season", "Style Number", "Color Code", "Color", "Name", "Wholesale (EUR)", 
                       "M.S.R.P. (EUR)", "Division", "Department", "Category", "Subcategory", 
                       "Product Notes", "Ship Start", "Ship End", "Prebook", "Country of Origin", 
                       "Fabric Description", "Total Price (EUR)", "Total Units"]
    
    ordered_main_cols = [col for col in main_cols_order if col in df_final.columns]  # Filtra le colonne disponibili
    additional_cols = [col for col in non_pivot_cols if col not in main_cols_order and col in df_final.columns]

    # Organizzare le colonne: colonne principali ordinate, seguite dalle taglie, poi le colonne extra
    final_columns = ordered_main_cols + additional_cols + final_size_order
    df_final = df_final[final_columns]

    return df_final

def convert_df_to_excel(df):
    """Converti il DataFrame in un oggetto Excel e restituisci il buffer."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Ottieni il foglio di lavoro
        worksheet = writer.sheets['Sheet1']

        # Imposta il formato per le colonne contenenti le date (se presenti)
        date_columns = ['Ship Start', 'Ship End']  # Aggiungi altre colonne data se necessario
        date_format = 'YYYY-MM-DD'  # Formato data desiderato
        for col in date_columns:
            if col in df.columns:
                # Trova l'indice della colonna
                col_index = df.columns.get_loc(col) + 1  # +1 perché l'indice di colonna in Excel parte da 1
                # Imposta il formato della data per la colonna
                worksheet.column_dimensions[chr(64 + col_index)].number_format = date_format

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
