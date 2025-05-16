import streamlit as st
from utils.interacao_db import db_config,load_and_prepare_data,create_user
import os
if not st.session_state.get('login_status', False) or not st.session_state.get('admin',False):
    st.warning('Página para acompanhamento interno')
    st.stop()
from PIL import Image
#Achando o caminho do icone
icon_path = os.path.join('.streamlit','Logo_azul_quadrada.PNG')
st.set_page_config(page_title="Gestão de Usuários", layout="wide",page_icon=icon_path)
#Achando o caminho da logo
img_path = os.path.join('.streamlit','Logo_goener_colorida.png')
img = Image.open(img_path)
# Redimensionar a imagem (alterando a altura e mantendo a proporção)
img = img.resize((int(img.width * (50 / img.height)), 50))  # Altura = 30 pixels
# Exibir a imagem redimensionada
st.image(img)
st.title('Gestão de Usuários')
#Carregar dados de usuários
users_df = load_and_prepare_data(db_config,'SELECT * FROM usuarios')
#Cadastro de Usuários
st.subheader('Cadastrar Usuário')
parceiro = st.text_input('Usuário')
email_parceiro = st.text_input('Email')
senha = st.text_input('Senha',type="password")
admin = st.checkbox('Admistrador')
infos_parceiro = {
    'username':parceiro,
    'email':email_parceiro,
    'senha':senha,
    'admin':admin
}
if st.button('Cadastrar usuário'):
    create_user(db_config,infos_parceiro)
    st.success('Usuário cadastrado com sucesso!')

