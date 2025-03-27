import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import base64
from datetime import datetime
import xlsxwriter
from io import BytesIO
import re
from streamlit import session_state as state

# Pagina-configuratie
st.set_page_config(
    page_title="Excel Bestandsvergelijker",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS voor een betere gebruikersinterface
st.markdown("""
<style>
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    .highlight {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
    }
    .diff-table td {
        text-align: left;
    }
    .success {
        color: green;
    }
    .warning {
        color: orange;
    }
    .danger {
        color: red;
    }
</style>
""", unsafe_allow_html=True)

# Functies voor bestandsverwerking
def read_excel_file(uploaded_file):
    """Excel bestand inlezen en beschikbare werkbladen ophalen"""
    if uploaded_file is not None:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            return excel_file, sheet_names
        except Exception as e:
            st.error(f"Fout bij het inlezen van het bestand: {e}")
            return None, []
    return None, []

def load_sheet_data(excel_file, sheet_name):
    """Gegevens van een specifiek werkblad inladen"""
    try:
        df = excel_file.parse(sheet_name)
        return df
    except Exception as e:
        st.error(f"Fout bij het inladen van werkblad {sheet_name}: {e}")
        return None

def filter_dataframe(df, column, filter_value, operation='equals'):
    """DataFrame filteren op basis van kolom en waarde"""
    if df is None or column not in df.columns:
        return df
    
    if operation == 'equals':
        return df[df[column] == filter_value]
    elif operation == 'contains':
        return df[df[column].astype(str).str.contains(filter_value, na=False)]
    elif operation == 'greater_than':
        return df[df[column] > filter_value]
    elif operation == 'less_than':
        return df[df[column] < filter_value]
    else:
        return df

def compare_values(val1, val2):
    """Vergelijk twee waarden met type-conversie en NaN afhandeling"""
    # Controleer op NaN waarden
    if pd.isna(val1) and pd.isna(val2):
        return True
    elif pd.isna(val1) or pd.isna(val2):
        return False
    
    # Converteer naar string voor vergelijking
    try:
        if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
            # Voor numerieke waarden, vergelijk met een kleine tolerantie
            return abs(float(val1) - float(val2)) < 1e-10
        elif isinstance(val1, datetime) and isinstance(val2, datetime):
            # Voor datetime objecten
            return val1 == val2
        else:
            # Andere types, converteer naar string
            return str(val1) == str(val2)
    except:
        # Bij fouten, gebruik gewone string vergelijking
        return str(val1) == str(val2)

def is_nan_or_inf(value):
    """Controleert of een waarde NaN of INF is"""
    if isinstance(value, (int, float)):
        return np.isnan(value) or np.isinf(value)
    return False

def is_nat(value):
    """Controleert of een waarde een pandas NaT (Not a Time) is"""
    if pd.api.types.is_datetime64_any_dtype(type(value)) or isinstance(value, pd.Timestamp):
        return pd.isna(value)
    return False

def is_problematic_value(value):
    """Controleert of een waarde problematisch is voor Excel export (NaN, INF, NaT)"""
    if is_nan_or_inf(value):
        return True
    if is_nat(value):
        return True
    return False

def export_changes_to_excel(diff_data, key_column, filename="veranderde_records.xlsx"):
    """
    Exporteer veranderde records naar een Excel-bestand met rode markering voor veranderingen
    
    Parameters:
    - diff_data: Dictionary met verschillen per sleutel
    - key_column: Naam van de sleutelkolom
    - filename: Naam van het bestand om naar te exporteren
    
    Returns:
    - BytesIO object met het Excel bestand
    """
    changes_data = []
    all_columns = set()
    
    for key, result in diff_data.items():
        if len(result['differences']) > 0:
            for diff in result['differences']:
                # Converteer timestamps naar strings
                old_value = str(diff['old_value']) if diff['old_value'] is not None else ''
                new_value = str(diff['new_value']) if diff['new_value'] is not None else ''
                
                changes_data.append({
                    'ID': key,
                    'Kolom': diff['column'],
                    'Oude Waarde': old_value,
                    'Nieuwe Waarde': new_value
                })
                all_columns.add(diff['column'])

    if changes_data:
        df = pd.DataFrame(changes_data)
        buffer = BytesIO()
        
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Veranderde Records', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Veranderde Records']
            
            # Stel standaard celformaat in
            standard_format = workbook.add_format({'text_wrap': True})
            
            # Pas automatische kolombreedte toe
            for idx, col in enumerate(df.columns):
                max_length = max(
                    len(str(col)),
                    df[col].astype(str).apply(len).max()
                )
                worksheet.set_column(idx, idx, max_length + 2, standard_format)
            
            # Header opmaak
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D9D9D9',
                'border': 1
            })
            
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

        # Belangrijk: zorg dat alle data geschreven is voordat we de buffer teruggeven
        buffer.seek(0)
        return buffer
    else:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            pd.DataFrame().to_excel(writer, sheet_name='Geen Veranderingen')
        buffer.seek(0)
        return buffer

def get_download_link_for_bytes(bytes_data, filename, text):
    """Creëer een downloadlink voor een BytesIO object"""
    b64 = base64.b64encode(bytes_data.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">{text}</a>'
    return href

def handle_duplicates(df, key_column):
    """
    Behandelt duplicaten in een DataFrame voor een specifieke kolom.
    Returns: DataFrame zonder duplicaten en aantal verwijderde duplicaten
    """
    # Tel originele rijen
    original_count = len(df)
    
    # Identificeer duplicaten
    duplicates = df[df[key_column].duplicated(keep='first')]
    duplicate_count = len(duplicates)
    
    if duplicate_count > 0:
        # Behoud alleen de eerste occurrence van elke waarde
        df = df.drop_duplicates(subset=[key_column], keep='first')
        
    return df, duplicate_count

# Hoofdfunctie voor de app
def main():
    st.title("📊 Excel Bestandsvergelijker - Focus op Veranderingen")
    st.markdown("Upload twee Excel-bestanden om veranderingen tussen gekoppelde records te analyseren.")
    
    # Sidebar voor bestandsupload
    with st.sidebar:
        st.header("Bestandsupload")
        
        # Nieuw bestand uploaden
        st.subheader("Nieuw Bestand")
        new_file = st.file_uploader("Upload het nieuwe Excel-bestand", type=['xlsx', 'xls'])
        
        # Oud bestand uploaden
        st.subheader("Oud Bestand")
        old_file = st.file_uploader("Upload het oude Excel-bestand", type=['xlsx', 'xls'])
    
    # Hoofdinhoud
    if new_file is not None and old_file is not None:
        st.header("1. Bestandsinformatie")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Nieuw Bestand")
            new_excel, new_sheets = read_excel_file(new_file)
            st.write(f"**Bestandsnaam:** {new_file.name}")
            st.write(f"**Werkbladen:** {len(new_sheets)}")
            
            # Werkblad selecteren
            new_sheet = st.selectbox("Selecteer werkblad (nieuw bestand):", new_sheets)
            
            # Werkblad inladen
            if new_sheet:
                new_df = load_sheet_data(new_excel, new_sheet)
                st.write(f"**Aantal rijen:** {len(new_df)}")
                st.write(f"**Aantal kolommen:** {len(new_df.columns)}")
                
                with st.expander("Bekijk kolomnamen"):
                    st.write(", ".join(new_df.columns.tolist()))
        
        with col2:
            st.subheader("Oud Bestand")
            old_excel, old_sheets = read_excel_file(old_file)
            st.write(f"**Bestandsnaam:** {old_file.name}")
            st.write(f"**Werkbladen:** {len(old_sheets)}")
            
            # Werkblad selecteren
            old_sheet = st.selectbox("Selecteer werkblad (oud bestand):", old_sheets)
            
            # Werkblad inladen
            if old_sheet:
                old_df = load_sheet_data(old_excel, old_sheet)
                st.write(f"**Aantal rijen:** {len(old_df)}")
                st.write(f"**Aantal kolommen:** {len(old_df.columns)}")
                
                with st.expander("Bekijk kolomnamen"):
                    st.write(", ".join(old_df.columns.tolist()))
        
        # Als beide bestanden geladen zijn, ga verder met de analyse
        if 'new_df' in locals() and 'old_df' in locals():
            st.header("2. Voorvertoning van de Gegevens")
            
            tab1, tab2 = st.tabs(["Nieuw Bestand", "Oud Bestand"])
            
            with tab1:
                # Filtering toevoegen
                if len(new_df.columns) > 0:
                    filter_col = st.selectbox("Filter op kolom (optioneel):", ["Geen filter"] + list(new_df.columns), key="new_filter_col")
                    
                    if filter_col != "Geen filter":
                        filter_op = st.selectbox("Filter operatie:", ["equals", "contains", "greater_than", "less_than"], key="new_filter_op")
                        
                        if filter_op in ["equals", "contains"]:
                            filter_val = st.text_input("Filter waarde:", key="new_filter_val")
                        else:
                            try:
                                filter_val = st.number_input("Filter waarde:", key="new_filter_val_num")
                            except:
                                filter_val = st.text_input("Filter waarde:", key="new_filter_val_text")
                        
                        new_df_filtered = filter_dataframe(new_df, filter_col, filter_val, filter_op)
                        st.write(f"Gefilterde data: {len(new_df_filtered)} rijen")
                        st.dataframe(new_df_filtered.head(100))
                    else:
                        st.write(f"Data preview (eerste 100 rijen van {len(new_df)} rijen):")
                        st.dataframe(new_df.head(100))
            
            with tab2:
                # Filtering toevoegen
                if len(old_df.columns) > 0:
                    filter_col = st.selectbox("Filter op kolom (optioneel):", ["Geen filter"] + list(old_df.columns), key="old_filter_col")
                    
                    if filter_col != "Geen filter":
                        filter_op = st.selectbox("Filter operatie:", ["equals", "contains", "greater_than", "less_than"], key="old_filter_op")
                        
                        if filter_op in ["equals", "contains"]:
                            filter_val = st.text_input("Filter waarde:", key="old_filter_val")
                        else:
                            try:
                                filter_val = st.number_input("Filter waarde:", key="old_filter_val_num")
                            except:
                                filter_val = st.text_input("Filter waarde:", key="old_filter_val_text")
                        
                        old_df_filtered = filter_dataframe(old_df, filter_col, filter_val, filter_op)
                        st.write(f"Gefilterde data: {len(old_df_filtered)} rijen")
                        st.dataframe(old_df_filtered.head(100))
                    else:
                        st.write(f"Data preview (eerste 100 rijen van {len(old_df)} rijen):")
                        st.dataframe(old_df.head(100))
            
            st.header("3. Datasets Koppelen & Veranderingen Analyseren")
            
            # Gemeenschappelijke kolommen bepalen
            common_columns = list(set(new_df.columns).intersection(set(old_df.columns)))
            
            if len(common_columns) > 0:
                st.write(f"**Gemeenschappelijke kolommen:** {len(common_columns)}")
                
                # Sleutelkolom selecteren
                key_column = st.selectbox("Selecteer een kolom om de datasets op te verbinden:", common_columns)
                
                if key_column:
                    # Tel duplicaten in beide datasets
                    new_duplicates_count = new_df[key_column].duplicated().sum()
                    old_duplicates_count = old_df[key_column].duplicated().sum()
                    
                    if new_duplicates_count > 0 or old_duplicates_count > 0:
                        st.warning(f"""
                        Gevonden duplicaten in kolom '{key_column}':
                        - Nieuw bestand: {new_duplicates_count} duplicaten
                        - Oud bestand: {old_duplicates_count} duplicaten
                        """)
                        
                        # Toon voorbeeld van duplicaten in nieuw bestand
                        if new_duplicates_count > 0:
                            with st.expander(f"Toon duplicaten in nieuw bestand voor '{key_column}'"):
                                duplicate_values = new_df[new_df[key_column].duplicated(keep=False)].sort_values(key_column)
                                st.dataframe(duplicate_values)
                        
                        # Optie om duplicaten te verwijderen
                        remove_duplicates = st.radio(
                            "Wil je de duplicaten verwijderen voor de vergelijking?",
                            options=["Ja, verwijder duplicaten (behoud eerste waarde)", "Nee, behoud alle records"],
                            index=0
                        )
                        
                        if remove_duplicates == "Ja, verwijder duplicaten (behoud eerste waarde)":
                            new_df, new_dup_removed = handle_duplicates(new_df, key_column)
                            old_df, old_dup_removed = handle_duplicates(old_df, key_column)
                            st.success(f"""
                            Duplicaten verwijderd:
                            - Nieuw bestand: {new_dup_removed} records
                            - Oud bestand: {old_dup_removed} records
                            """)

                if st.button("Datasets Koppelen en Veranderingen Analyseren"):
                    with st.spinner("Bezig met koppelen en veranderingen analyseren..."):
                        # Handel duplicaten af in beide datasets
                        new_df, new_duplicates = handle_duplicates(new_df, key_column)
                        old_df, old_duplicates = handle_duplicates(old_df, key_column)
                        
                        if new_duplicates > 0:
                            st.warning(f"{new_duplicates} duplicaten verwijderd uit het nieuwe bestand voor kolom '{key_column}'. Eerste waarde van elke duplicaat is behouden.")
                        
                        if old_duplicates > 0:
                            st.warning(f"{old_duplicates} duplicaten verwijderd uit het oude bestand voor kolom '{key_column}'. Eerste waarde van elke duplicaat is behouden.")
                        
                        # Sleutelwaarden in beide datasets
                        new_keys = set(new_df[key_column])
                        old_keys = set(old_df[key_column])
                        
                        # Records-analyse
                        keys_only_in_new = new_keys - old_keys  # Nieuwe records
                        keys_only_in_old = old_keys - new_keys  # Verwijderde records  
                        common_keys = new_keys.intersection(old_keys)  # Gekoppelde records
                        
                        # Informatie over de koppeling
                        st.subheader("Records-analyse")
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.metric("Nieuwe records", len(keys_only_in_new), 
                                     f"{len(keys_only_in_new)/len(new_keys):.1%} van nieuwe dataset")
                            st.write("*Records die wel in het nieuwe bestand zitten maar niet in het oude*")
                        
                        with col2:
                            st.metric("Verwijderde records", len(keys_only_in_old), 
                                     f"{len(keys_only_in_old)/len(old_keys):.1%} van oude dataset")
                            st.write("*Records die wel in het oude bestand zitten maar niet in het nieuwe*")
                        
                        with col3:
                            st.metric("Gekoppelde records", len(common_keys), 
                                     f"{len(common_keys)/len(new_keys):.1%} van nieuwe dataset")
                            st.write("*Records die in beide bestanden voorkomen*")
                        
                        # Alleen vergelijking uitvoeren als er gekoppelde records zijn
                        if len(common_keys) > 0:
                            # Filter beide datasets op gemeenschappelijke sleutels
                            new_df_common = new_df[new_df[key_column].isin(common_keys)]
                            old_df_common = old_df[old_df[key_column].isin(common_keys)]
                            
                            # Set index voor snellere lookup
                            new_df_common = new_df_common.set_index(key_column)
                            old_df_common = old_df_common.set_index(key_column)
                            
                            # Gemeenschappelijke kolommen voor vergelijking
                            compare_columns = [col for col in common_columns if col != key_column]
                            
                            # Datastructuur voor de verschillen
                            diff_results = {}
                            column_diff_count = {col: 0 for col in compare_columns}
                            
                            # Voor elke gemeenschappelijke sleutel
                            for key in common_keys:
                                try:
                                    # Haal beide rijen op
                                    new_row = new_df_common.loc[key]
                                    old_row = old_df_common.loc[key]
                                    
                                    # Initialiseer voor deze key
                                    diff_results[key] = {
                                        'total_columns': len(compare_columns),
                                        'matching_columns': 0,
                                        'differences': []
                                    }
                                    
                                    # Vergelijk elke kolom
                                    for col in compare_columns:
                                        new_val = new_row[col] if col in new_row.index else None
                                        old_val = old_row[col] if col in old_row.index else None
                                        
                                        # Vergelijk de waarden
                                        if compare_values(new_val, old_val):
                                            diff_results[key]['matching_columns'] += 1
                                        else:
                                            diff_results[key]['differences'].append({
                                                'column': col,
                                                'old_value': old_val,
                                                'new_value': new_val
                                            })
                                            column_diff_count[col] += 1
                                except Exception as e:
                                    st.error(f"Fout bij het vergelijken van rij met sleutel {key}: {e}")
                                    continue
                            
                            # Aantal rijen met veranderingen
                            changed_rows = sum(1 for key, result in diff_results.items() if len(result['differences']) > 0)
                            unchanged_rows = len(common_keys) - changed_rows
                            
                            # Focus op veranderingen in gekoppelde records
                            st.subheader("Veranderingen in Gekoppelde Records")
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.metric("Veranderde records", changed_rows, 
                                         f"{changed_rows/len(common_keys):.1%} van gekoppelde records")
                            
                            with col2:
                                st.metric("Onveranderde records", unchanged_rows, 
                                         f"{unchanged_rows/len(common_keys):.1%} van gekoppelde records")
                            
                            # Tabel met aantal verschillen per kolom
                            if changed_rows > 0:
                                st.subheader("Verschillen per kolom")
                                diff_counts_df = pd.DataFrame({
                                    'Kolom': column_diff_count.keys(),
                                    'Aantal Verschillen': column_diff_count.values(),
                                    'Percentage': [count/len(common_keys)*100 for count in column_diff_count.values()]
                                }).sort_values('Aantal Verschillen', ascending=False)
                                
                                # Toon alleen kolommen met verschillen
                                diff_counts_df = diff_counts_df[diff_counts_df['Aantal Verschillen'] > 0]
                                
                                if not diff_counts_df.empty:
                                    # Maak een interactieve barchart
                                    fig = px.bar(
                                        diff_counts_df, 
                                        x='Kolom', 
                                        y='Aantal Verschillen',
                                        hover_data=['Percentage'],
                                        title='Aantal Verschillen per Kolom',
                                        color='Aantal Verschillen',
                                        color_continuous_scale='Reds'
                                    )
                                    st.plotly_chart(fig, use_container_width=True)
                                    
                                    # Tabel met de verschillen per kolom
                                    st.dataframe(diff_counts_df)
                                    
                                    # Exporteer veranderde records
                                    st.subheader("Exporteer Veranderde Records")
                                    st.write("Download een Excel-bestand met alle veranderde records. Veranderingen worden in rood gemarkeerd.")
                                    
                                    # Genereer het Excel-bestand
                                    excel_buffer = export_changes_to_excel(diff_results, key_column)
                                    export_filename = f"veranderde_records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                                    
                                    # Download knop met rerun
                                    if st.download_button(
                                        label="Download Veranderde Records (Excel)",
                                        data=excel_buffer,
                                        file_name=export_filename,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    ):
                                        st.rerun()
                                else:
                                    st.info("Geen verschillen gevonden tussen de gemeenschappelijke records.")
                            else:
                                st.success("Alle gekoppelde records zijn identiek. Er zijn geen veranderingen gevonden.")
                        else:
                            st.warning("Er zijn geen gemeenschappelijke records gevonden om te vergelijken.")
            else:
                st.error("De bestanden hebben geen gemeenschappelijke kolommen om te vergelijken.")

if __name__ == "__main__":
    main()
