# env neuralprophet

# import necessary libraries
import pandas as pd
import numpy as np
import warnings
import matplotlib.pyplot as plt
import plotly.express as px
warnings.filterwarnings('ignore')
from io import BytesIO
import streamlit as st
import plotly.graph_objects as go
import calendar

####### Impaginazione

st.set_page_config(layout="wide")

url_immagine = 'https://github.com/MarcelloGalimberti/Ducati_RCCP/blob/main/Ducati_red_logo-4.PNG?raw=true'

col_1, col_2 = st.columns([1, 5])

with col_1:
    st.image(url_immagine, width=200)

with col_2:
    st.title('RCCP 2026')

st.header('Caricamento dati da remoto | GitHub', divider='red')
st.write(
    'Dati caricati da file Excel presenti su GitHub: '
    '\n- abbinamento modello-famiglia '
    '\n- CDC Veicolo 2026 '
    '\n- CDC Motore 2026 '
    '\n- infine, i file PPP 2026 e Calendario RCCP 2026 devono essere caricato tramite uploader.')

####### Caricamento dati

url_CDC_veicolo_2026 = 'https://github.com/MarcelloGalimberti/Ducati_RCCP/blob/main/CDC%20Veicolo%202026.xlsx?raw=true'
url_CDC_motore_2026 = 'https://github.com/MarcelloGalimberti/Ducati_RCCP/blob/main/CDC%20Motore%202026.xlsx?raw=true'


# Caricamento dei file Excel con i dati delle famiglie, linee e mesi per veicolo e motore
df_famiglia_linea_mese_veicolo = pd.read_excel(url_CDC_veicolo_2026)#, parse_dates=True, index_col=0)
df_melt_veicolo = df_famiglia_linea_mese_veicolo.melt(id_vars='FAMIGLIA')
df_melt_veicolo.columns=['famiglia','anno-mese','linea'] 

df_famiglia_linea_mese_motore = pd.read_excel(url_CDC_motore_2026)#, parse_dates=True, index_col=0)
df_melt_motore = df_famiglia_linea_mese_motore.melt(id_vars='FAMIGLIA')
df_melt_motore.columns=['famiglia','anno-mese','linea']

# Unione dei DataFrame per veicolo e motore
df_melt = pd.concat([df_melt_veicolo, df_melt_motore], ignore_index=True)
# Rinomina le colonne per chiarezza
df_melt.columns = ['Famiglia', 'Anno-Mese', 'Linea']
# Anno-Mese in formato datetime anno - mese - giorno
df_melt['Anno-Mese'] = pd.to_datetime(df_melt['Anno-Mese'], format='%d/%m/%Y')
df_melt.rename(columns={'Anno-Mese': 'Data', 'Linea':'CDC'}, inplace=True)

# Visualizza il DataFrame unito
#st.write('Abbinamento famiglie, linee e mesi per veicolo e motore:')
#st.dataframe(df_melt, use_container_width=True)


url_abbinamento_modello_famiglia = 'https://github.com/MarcelloGalimberti/Ducati_RCCP/blob/main/abbinamento_modello_famiglia.xlsx?raw=true'

# Caricamento del file di abbinamento modello-famiglia
df_modello_famiglia = pd.read_excel(url_abbinamento_modello_famiglia)#, parse_dates=True, index_col=0)
#st.write('Abbinamento Modello - Famiglia:')
#st.dataframe(df_modello_famiglia, use_container_width=True)

# Caricamento del file PPP 2026
st.header('Caricamento PPP 2026', divider='red')
uploaded_files = st.file_uploader(
    "Carica PPP 2026 (PR74OP 2026 Bologna_esp 002_.xlsx)",
    #accept_multiple_files=True # caricarli tutti e poi selezionare 2026 (in un secondo momento, 2030 è diverso)
)
if not uploaded_files:
    st.stop()

# Caricamento del file Calendario PPP 2026
st.header('Caricamento Calendario 2026', divider='red')
uploaded_calendario = st.file_uploader(
    "Carica Calendario 2026 (Calendario_RCCP_2026.xlsx)",
    #accept_multiple_files=True # caricarli tutti e poi selezionare 2026 (in un secondo momento, 2030 è diverso)
)
if not uploaded_calendario:
    st.stop()



# Caricamento del file PPP 2026, specificando il nome del foglio e le righe da saltare
df_PPP = pd.read_excel(uploaded_files, sheet_name='PR74 OP 2026', skiprows=[1, 2, 3, 4, 5, 6, 7, 8, 9], header=0)
# Elimina le colonne non necessarie
df_PPP = df_PPP.drop(columns=['I°sem', 'II°sem', 'TOT'])
# Prendo il nome della prima colonna di ciascun DataFrame
colonna_chiave_ppp = df_PPP.columns[0]
colonna_chiave_modello = df_modello_famiglia.columns[0]
# Filtro df_PPP mantenendo solo i valori presenti in df_modello_famiglia
df_PPP = df_PPP[df_PPP[colonna_chiave_ppp].isin(df_modello_famiglia[colonna_chiave_modello])].fillna(0)
# Genera 12 date mensili a partire da gennaio 2026
date_nuove = pd.date_range(start="2026-01-01", periods=12, freq='MS').strftime('%d/%m/%Y')
# Costruisce la lista completa: prima colonna invariata + nuove date
df_PPP.columns = [df_PPP.columns[0]] + list(date_nuove)
# Rinomina la prima colonna in 'Modello'
df_PPP.rename(columns={df_PPP.columns[0]: 'Modello'}, inplace=True)
# Crea un dizionario Modello → Famiglia
mappa_famiglia = df_modello_famiglia.set_index('Modello')['Famiglia']
# Aggiunge la colonna Famiglia a df_PPP usando la mappatura
df_PPP['Famiglia'] = df_PPP['Modello'].map(mappa_famiglia)
#st.write('PPP 2026 processato:')
#st.dataframe(df_PPP, use_container_width=True)

# 1. Determina le colonne delle date (escludendo 'Modello' e 'Famiglia')
colonne_mensili = df_PPP.columns[1:-1]  # da colonna 1 a 12
# 2. Raggruppa per Famiglia e somma i valori mensili
df_sommato_per_famiglia = df_PPP.groupby('Famiglia')[colonne_mensili].sum()
# Elimina le righe dove tutti i valori sono zero
df_sommato_per_famiglia = df_sommato_per_famiglia[(df_sommato_per_famiglia != 0).any(axis=1)]
#st.dataframe(df_sommato_per_famiglia, use_container_width=True)
df_unpivot = df_sommato_per_famiglia.reset_index().melt(
    id_vars='Famiglia',
    var_name='Data',
    value_name='Qty'
)
df_unpivot['Data'] = pd.to_datetime(df_unpivot['Data'], format='%d/%m/%Y')
st.write('PPP processato: Quantità per Famiglia e Data:')
st.dataframe(df_unpivot, use_container_width=True)

# volumi per mese
df_volumi_per_mese = df_unpivot.groupby('Data')['Qty'].sum().reset_index()
#st.write('Volumi totali per mese:')
#st.dataframe(df_volumi_per_mese, use_container_width=True)

# Grafico
# Crea la colonna per l'asse x con il formato "mese-anno" (in italiano)
mesi_it = [
    '', 'gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno',
    'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'
]
df_volumi_per_mese['Mese-Anno'] = df_volumi_per_mese['Data'].apply(
    lambda x: f"{mesi_it[x.month]}-{x.year}"
)

# Ordina per data
df_volumi_per_mese = df_volumi_per_mese.sort_values('Data')

st.header('Dati RCCP 2026', divider='red')
# Grafico
fig = px.bar(
    df_volumi_per_mese,
    x='Mese-Anno',
    y='Qty',
    title='Volumi mensili 2026',
    labels={'Qty': 'Qty', 'Mese-Anno': 'Mese-Anno'}
)
fig.update_layout(xaxis_tickangle=45, yaxis_title="Qty", xaxis_title="Mese-Anno")

st.plotly_chart(fig, use_container_width=True)


###########
# Caricamento del file Calendario RCCP 2026

#url_calendario_RCCP_2026 = 'https://github.com/MarcelloGalimberti/Ducati_RCCP/blob/main/Calendario_RCCP_2026.xlsx?raw=true'

df_TC = pd.read_excel(uploaded_calendario, sheet_name='DB_Risorse Mensile', parse_dates=True)
df_TC.rename(columns={'FAMIGLIA': 'Famiglia', 'MESE': 'Data'}, inplace=True)

# Merge con volumi
df_RCCP = pd.merge(
    df_TC,
    df_unpivot,
    left_on=['Famiglia', 'Data'],
    right_on=['Famiglia', 'Data'],
    how='left'
)

#st.write('Calendario RCCP 2026 con volumi:')
#st.dataframe(df_RCCP, use_container_width=True)

# Merge linee veicolo e motore con Qty
df_linee = pd.merge(
    df_melt,
    df_unpivot,
    left_on=['Famiglia', 'Data'],
    right_on=['Famiglia', 'Data'],
    how='left'
)

#st.write('Linee veicolo e motore con volumi:')
#st.dataframe(df_linee, use_container_width=True)

df_aggregato = df_linee.groupby(['Data', 'CDC'])['Qty'].sum().reset_index()

#st.write('Linee aggregate per CDC e Data:')
#st.dataframe(df_aggregato, use_container_width=True)

# Filtra df_RCCP per mantenere solo le linee presenti in df_linee
df_RCCP_mask = df_RCCP['CDC'].isin(df_aggregato['CDC'])

#st.write('RCCP filtrato per CdC presenti in df_linee:')
#st.dataframe(df_RCCP[df_RCCP_mask], use_container_width=True)

#Fai un merge per estrarre i Qty corretti da df_linee
df_merge = df_RCCP[df_RCCP_mask].merge(
    df_aggregato[['Data', 'CDC', 'Qty']],
    on=['Data', 'CDC'],
    how='left',
    suffixes=('', '_from_linee')
)

#st.write('Merge tra RCCP e linee per ottenere Qty:')
#st.dataframe(df_merge, use_container_width=True)

# Fai un merge per unire i dati da df_merge su Data e CDC
df_temp = df_RCCP.merge(
    df_merge[['Data', 'CDC', 'Qty_from_linee']],
    on=['Data', 'CDC'],
    how='left'
)

# Sostituisci i valori di Qty solo dove Qty_from_line non è NaN
df_temp['Qty'] = df_temp['Qty_from_linee'].combine_first(df_temp['Qty'])

# Rimuovi la colonna temporanea
df_RCCP = df_temp.drop(columns='Qty_from_linee')

########
# Filtra solo i CDC di interesse
cdc_target = [571, 572]
maschera_cdc = df_RCCP['CDC'].isin(cdc_target)

# Merge su 'Data' per ottenere i nuovi Qty
df_merge = df_RCCP.merge(
    df_volumi_per_mese[['Data', 'Qty']],
    on='Data',
    how='left',
    suffixes=('', '_nuovo')
)

# Sostituzione condizionata: CDC in [571,572] e Qty_nuovo disponibile
df_merge['Qty'] = df_merge.apply(
    lambda row: row['Qty_nuovo'] if row['CDC'] in cdc_target and pd.notnull(row['Qty_nuovo']) else row['Qty'],
    axis=1
)

# Elimina la colonna temporanea
df_RCCP = df_merge.drop(columns='Qty_nuovo')

# Aggiunge famiglie

# 1. Raggruppa df_melt per Data e CDC, concatenando Famiglie
df_famiglie_concat = (
    df_melt
    .dropna(subset=['Famiglia'])
    .groupby(['Data', 'CDC'])['Famiglia']
    .apply(lambda x: ', '.join(sorted(set(map(str, x)))))
    .reset_index()
)

# 2. Crea dizionario (Data, CDC) → stringa concatenata
famiglia_dict = {
    (row['Data'], row['CDC']): row['Famiglia']
    for _, row in df_famiglie_concat.iterrows()
}

# 3. Funzione di aggiornamento condizionato
def aggiorna_famiglia(row):
    if row['CDC'] in [571, 572]:
        return "Tutte le famiglie"
    key = (row['Data'], row['CDC'])
    return famiglia_dict.get(key, row['Famiglia'])

# 4. Applica la funzione al dataframe
df_RCCP['Famiglia'] = df_RCCP.apply(aggiorna_famiglia, axis=1)


#st.write('RCCP con Qty aggiornati da df_linee e banchi riparazione:')
#st.dataframe(df_RCCP, use_container_width=True)


#########
# Calcolo del workload
df_RCCP['Workload'] = df_RCCP['Qty'] * df_RCCP['T.C.']/60/df_RCCP['OEE']

# Calcolo della saturazione
# Tabella pivot
pivot_df = pd.pivot_table(
    df_RCCP,
    index=['Data', 'REPARTO', 'CDC', 'RISORSA PRIMARIA'],
    values=["MOLTEPLICITA'", 'ore/mese', 'Workload'],
    aggfunc={
        "MOLTEPLICITA'": 'mean',
        'ore/mese': 'mean',
        'Workload': 'sum'
    }
).reset_index()

pivot_df = pivot_df.round(2)

# 1. Calcola la Capacity = MOLTEPLICITA' × ore/mese
pivot_df['Capacity'] = pivot_df["MOLTEPLICITA'"] * pivot_df['ore/mese']

# 2. Calcola la Saturazione = Workload / Capacity
# Usa np.where per evitare divisioni per zero
import numpy as np

pivot_df['Saturazione'] = np.where(
    pivot_df['Capacity'] > 0,
    pivot_df['Workload'] / pivot_df['Capacity'],
    np.nan
)

# 3. Formatta la Saturazione come percentuale con un decimale
pivot_df['Saturazione'] = (pivot_df['Saturazione'] * 100).round(1).astype(str) + '%'


st.subheader('Tabella Saturazione per Data, Reparto, CDC e Risorsa Primaria:', divider='red')
st.dataframe(pivot_df, use_container_width=True)    

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Foglio1')
    return output.getvalue()

# Crea il bottone per scaricare file Saturazione
saturazione_file = to_excel_bytes(pivot_df)
st.download_button(
    label="📥 Scarica file Excel saturazione",
    data=saturazione_file,
    file_name='df_saturazione.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)



#########
# Step 1: Calcola Saturazione riga per riga
df_RCCP['Capacity'] = df_RCCP["MOLTEPLICITA'"] * df_RCCP['ore/mese']

# Evita divisioni per 0
df_RCCP['Saturazione'] = np.where(
    df_RCCP['Capacity'] > 0,
    df_RCCP['Workload'] / df_RCCP['Capacity'],
    np.nan
)


# Step 2: Crea pivot con Famiglia come colonne
pivot_famiglia = pd.pivot_table(
    df_RCCP,
    index=['Data', 'REPARTO', 'CDC', 'RISORSA PRIMARIA'],
    columns='Famiglia',
    values='Saturazione',
    aggfunc='mean'  # se ci sono più righe per combinazione
).reset_index()


# Seleziona solo le colonne delle Famiglie (dopo reset_index)
fam_cols = pivot_famiglia.columns.difference(['Data', 'REPARTO', 'CDC', 'RISORSA PRIMARIA'])

# Moltiplica per 100 e formatta
pivot_famiglia[fam_cols] = pivot_famiglia[fam_cols] \
    .multiply(100) \
    .round(1) \
    .astype(str) + '%'

#########


# 1. (Solo se devi ricreare il df da dict, altrimenti salta)
# pivot_famiglia = pd.DataFrame(dati) # <-- solo se il df non esiste già

# 2. Colonne chiave
indici = ['Data', 'REPARTO', 'CDC', 'RISORSA PRIMARIA', 'Saturazione']
fam_cols = [c for c in pivot_famiglia.columns if c not in indici]

# 3. Funzione robusta per convertire le stringhe percentuali in float
def clean_pct(val):
    if isinstance(val, str):
        v = val.strip().replace('%','').replace(',','.')
        if v.lower() in ('nan',''):
            return np.nan
        try:
            return float(v)
        except Exception:
            return np.nan
    if pd.isna(val):
        return np.nan
    return val

# 4. Applica la funzione su ogni valore delle colonne Famiglia
df_numeric = pivot_famiglia[fam_cols].applymap(clean_pct)

# 5. Per ogni riga: se TUTTE le colonne famiglia sono nan, la saturazione sarà vuota, altrimenti la somma formattata
saturazione_valori = df_numeric.sum(axis=1, skipna=True)
tutti_nan = df_numeric.isna().all(axis=1)

# 6. Colonna finale
pivot_famiglia['Saturazione'] = [
    "" if is_nan else f"{val:.1f}%"
    for is_nan, val in zip(tutti_nan, saturazione_valori)
]

# 7. (Opzionale) Rimuovi visualmente 'nan%' dalle colonne famiglia
pivot_famiglia[fam_cols] = pivot_famiglia[fam_cols].replace('nan%', '', regex=False)

st.subheader('Tabella di Saturazione per Famiglia:', divider='red')
st.dataframe(pivot_famiglia, use_container_width=True)



# Crea il bottone per scaricare file
#RCCP_file = to_excel_bytes(df_RCCP)
#st.download_button(
#    label="📥 Scarica file RCCP",
#    data=RCCP_file,
#    file_name='df_RCCP.xlsx',
#    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
#)



# Crea il bottone per scaricare file Saturazione per Famiglia
saturazione_famiglia_file = to_excel_bytes(pivot_famiglia)
st.download_button(
    label="📥 Scarica file Excel saturazione per famiglia",
    data=saturazione_famiglia_file,
    file_name='df_saturazione_famiglia.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

############# Grafici 1

st.subheader('Analisi saturazione mensile:', divider='red')

# Funzione per il titolo data
def mese_anno(dt):
    # Se dt è un Timestamp, estrai mese e anno
    if isinstance(dt, pd.Timestamp):
        return f"{calendar.month_name[dt.month]} {dt.year}"
    # Se è una stringa, tenta conversione
    try:
        dt = pd.to_datetime(dt)
        return f"{calendar.month_name[dt.month]} {dt.year}"
    except Exception:
        return str(dt)

# Prepara la colonna x_label
pivot_famiglia['x_label'] = (
    pivot_famiglia['REPARTO'].astype(str) + ' | ' +
    pivot_famiglia['CDC'].astype(str) + ' | ' +
    pivot_famiglia['RISORSA PRIMARIA'].astype(str)
)

# Trova le colonne delle famiglie
indici = ['Data', 'REPARTO', 'CDC', 'RISORSA PRIMARIA', 'Saturazione', 'x_label']
fam_cols = [c for c in pivot_famiglia.columns if c not in indici]

for dt in pivot_famiglia['Data'].unique():
    df_dt = pivot_famiglia[pivot_famiglia['Data'] == dt]
    fig = go.Figure()
    for fam in fam_cols:
        y_val = df_dt[fam].replace('nan%', '', regex=False).replace('', np.nan)
        y_val = y_val.apply(lambda v: float(str(v).replace('%','').replace(',','.')) if pd.notnull(v) and str(v).strip() != '' else 0)
        y_val = y_val.where(y_val > 0, 0)
        if y_val.sum() > 0:
            fig.add_bar(
                x=df_dt['x_label'],
                y=y_val,
                name=fam
            )
    # Aggiungi linea orizzontale al 100%
    fig.add_shape(
        type="line",
        x0=-0.5,
        y0=100,
        x1=len(df_dt['x_label']) - 0.5,
        y1=100,
        line=dict(color="red", width=2, dash="dash"),
    )
    fig.update_layout(
        barmode='stack',
        title=f"Saturazione per CDC e Famiglia<br><sup>{mese_anno(dt)}</sup>",
        xaxis_title="Reparto | CDC | Risorsa primaria",
        yaxis_title="Saturazione (%)",
        legend_title="Famiglia",
        yaxis=dict(ticksuffix='%', range=[0, max(110, df_dt[fam_cols].replace('nan%', '', regex=False)
                                                      .replace('', np.nan)
                                                      .applymap(lambda v: float(str(v).replace('%','').replace(',','.')) if pd.notnull(v) and str(v).strip() != '' else 0)
                                                      .max().max() + 10)]
        ),
        xaxis_tickangle=45,
        height=600
    )
    st.plotly_chart(fig, use_container_width=True)


st.stop()








veicolo_file = to_excel_bytes(df_melt_veicolo)
st.download_button(
    label="📥 Scarica file veicolo",
    data=veicolo_file,
    file_name='df_veicolo.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)