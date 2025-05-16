import streamlit as st
import os
from PIL import Image
import io
import datetime
#Checar se o usuário tem permissão de acesso
if not st.session_state.get('login_status', False):
    st.warning('Por favor, faça login.')
    st.stop()
#Importações internas
from utils.edit_powerpoint import powerpoint_edit
from utils.interacao_db import insert_proposal,db_config
#Achando o caminho do icone
icon_path = os.path.join('.streamlit','Logo_azul_quadrada.PNG')
st.set_page_config(page_title="Proposta Comercial", layout="wide",page_icon=icon_path)
#Achando o caminho da logo
img_path = os.path.join('.streamlit','Logo_goener_colorida.png')
img = Image.open(img_path)
# Redimensionar a imagem (alterando a altura e mantendo a proporção)
img = img.resize((int(img.width * (50 / img.height)), 50))  # Altura = 30 pixels
# Exibir a imagem redimensionada
st.image(img)
st.title('Proposta Comercial')
#Começo dos inputs do usuário
nome_cliente = st.text_input('Nome do cliente')

# Criando duas colunas para os inputs numéricos
col1, col2 = st.columns(2)

with col1:
    desconto = st.number_input('Desconto',step=5)
    custo_disponibilidade = st.selectbox('Custo de Disponibilidade',['Monofásico', 'Bifásico', 'Trifásico'])
    cip = st.number_input('Custo pago para iluminação pública', step=10)

with col2:
    consumo = st.number_input('Consumo (Kwh)',step=100)
    n_ucs = st.number_input("Número de UCs",step=1)

# Salvando as informações em um dicionário
infos = {
    'parceiro':st.session_state['current_user'],
    'cliente': nome_cliente,
    'desconto': desconto,
    'consumo': consumo,
    'custo_disponibilidade': custo_disponibilidade,
    'n_ucs': n_ucs,
    'horario':"datetime64[ns]",
    'cip': cip
}

create_button = st.button('Criar Apresentação')

if create_button:
    #Abertura do arquivo
    CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.path.dirname(CURRENT_DIR)
    prs_path = os.path.join(BASE_DIR,'.streamlit','APRESENTACAO_GD.pptx')
    with open(prs_path,'rb') as f:
        pptx_bytes = f.read()

    pptx_buffer = io.BytesIO(pptx_bytes)
    pptx_file = powerpoint_edit(infos,pptx_buffer)
    st.success('Apresentação Criada com Sucesso!')
    infos['horario'] = datetime.datetime.now().replace(microsecond=0)
    insert_proposal(db_config, infos)
    #Disponibilizar powerpoint para Download
    download_proposal = st.download_button(
        label="Baixar Apresentação Power Point",
        data=pptx_file,
        file_name=f"Apresentacao_{infos['cliente']}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )



