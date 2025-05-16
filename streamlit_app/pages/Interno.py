import pandas as pd
import streamlit as st
import plotly.express as px
import os
if not st.session_state.get('login_status', False) or not st.session_state.get('admin',False):
    st.warning('Página para acompanhamento interno')
    st.stop()
from PIL import Image
#Importação Interna
from utils.interacao_db import db_config,load_and_prepare_data
#Achando o caminho do icone
icon_path = os.path.join('.streamlit','Logo_azul_quadrada.PNG')
st.set_page_config(page_title="Acompanhamento", layout="wide",page_icon=icon_path)
#Achando o caminho da logo
img_path = os.path.join('.streamlit','Logo_goener_colorida.png')
img = Image.open(img_path)
# Redimensionar a imagem (alterando a altura e mantendo a proporção)
img = img.resize((int(img.width * (50 / img.height)), 50))  # Altura = 30 pixels
# Exibir a imagem redimensionada
st.image(img)
st.title('Acompanhamento de propostas')
st.subheader('Propostas')
#Carregar dataframe
df_propostas = load_and_prepare_data(db_config,'SELECT * FROM propostas')
df_propostas['horario'] = pd.to_datetime(df_propostas['horario']).dt.date
st.dataframe(df_propostas)
# Contar as ocorrências de cada parceiro
df_counts = df_propostas['parceiro'].value_counts().reset_index()
df_counts.columns = ['Parceiro', 'Propostas']
fig1 = px.bar(df_counts,x='Parceiro',y='Propostas')
fig1.update_traces(marker=dict(color='#009F98'))
st.plotly_chart(fig1)
#Série temporal
df_temporal = df_propostas.groupby(['parceiro','horario']).size().reset_index()
df_temporal.columns = ['Parceiro','Horario','Propostas']
all_pairs = pd.MultiIndex.from_product([df_temporal['Horario'].unique(), df_temporal['Parceiro'].unique()], names=['Horario', 'Parceiro'])

# Reindexar o DataFrame para incluir todas as combinações possíveis
df_temporal_full = df_temporal.set_index(['Horario', 'Parceiro']).reindex(all_pairs, fill_value=0).reset_index().sort_values(by='Horario')
fig2 = px.line(df_temporal_full,x='Horario',y='Propostas',color = 'Parceiro',markers = True)
st.plotly_chart(fig2)