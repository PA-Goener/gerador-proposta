import streamlit as st
from PIL import Image
import os
#Importações internas
from utils.auth import init_session_state
from utils.interacao_db import load_and_prepare_data,db_config

#Achando o caminho do icone
icon_path = os.path.join('.streamlit','Logo_azul_quadrada.PNG')
st.set_page_config(page_title="Página Inicial", layout="wide",page_icon=icon_path)
#Achando o caminho da logo
img_path = os.path.join('.streamlit','Logo_goener_colorida.png')
img = Image.open(img_path)
# Redimensionar a imagem (alterando a altura e mantendo a proporção)
img = img.resize((int(img.width * (50 / img.height)), 50))  # Altura = 30 pixels
# Exibir a imagem redimensionada
st.image(img)

def login():
    username = st.text_input("Usuário")
    password = st.text_input("Senha", type='password')

    if st.button("Login"):
        users_df = load_and_prepare_data(db_config,'SELECT * FROM usuarios')
        if username in users_df['username'].unique() and password == users_df[users_df['username']==username]['senha'].values[0]:
            st.session_state['login_status'] = True
            st.session_state['current_user'] = username
            st.session_state['admin'] = users_df[users_df['username'] == username]['admin'].values[0]
            st.success("Login realizado com sucesso!")
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos")

def main():
    init_session_state()

    if not st.session_state['login_status']:
        st.title("Bem-vindo!")
        login()
    else:
        st.title(f"Bem-vindo, {st.session_state['current_user']}!")
        if st.sidebar.button("Logout"):
            st.session_state['login_status'] = False
            st.session_state['current_user'] = None
            st.session_state['admin'] = False
            st.rerun()

if __name__ == '__main__':
    main()