import pandas as pd
from sqlalchemy import create_engine,text
import streamlit as st
import psycopg2

db_config = {
        'host': st.secrets["database"]["host"],
        'port': st.secrets["database"]["port"],
        'dbname': st.secrets["database"]["database"],
        'user': st.secrets["database"]["user"],
        'password': st.secrets["database"]["password"]
}

def load_and_prepare_data(db_config: dict, query: str) -> pd.DataFrame:
    connection_string = (
        f"postgresql://{db_config['user']}:{db_config['password']}@"
        f"{db_config['host']}:{db_config['port']}/{db_config['dbname']}"
    )
    engine = create_engine(connection_string)

    df = pd.read_sql_query(query, engine)

    return df


def insert_proposal(db_config,infos):
    connection_string = (
        f"postgresql://{db_config['user']}:{db_config['password']}@"
        f"{db_config['host']}:{db_config['port']}/{db_config['dbname']}"
    )

    engine = create_engine(connection_string)

    query = text("""
    INSERT INTO propostas(parceiro,horario,cliente,desconto,consumo,custo_disponibilidade,n_ucs,cip)
    VALUES(:parceiro,:horario,:cliente,:desconto,:consumo,:custo_disponibilidade,:n_ucs,:cip)
    """)
    with engine.connect() as connection:
        connection.execute(query, infos)
        connection.commit()

def create_user(db_config,partner_info):
    connection_string = (
        f"postgresql://{db_config['user']}:{db_config['password']}@"
        f"{db_config['host']}:{db_config['port']}/{db_config['dbname']}"
    )

    engine = create_engine(connection_string)

    query = text("""
        INSERT INTO usuarios(email,senha,username,admin)
        VALUES(:email,:senha,:username,:admin)
        """)
    with engine.connect() as connection:
        connection.execute(query, partner_info)
        connection.commit()