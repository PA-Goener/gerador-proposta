import streamlit as st

def init_session_state():
    if 'login_status' not in st.session_state:
        st.session_state['login_status'] = False
    if 'current_user' not in st.session_state:
        st.session_state['current_user'] = None



