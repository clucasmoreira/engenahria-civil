import streamlit as st
import tempfile
import os

from lista_mestra_web import gerar_lista_mestra

st.title("Lista Mestra de Projetos")

arquivos = st.file_uploader(
    "Selecione PDFs e TXTs",
    accept_multiple_files=True,
    type=["pdf", "txt"]
)

if arquivos:

    if st.button("Gerar Lista Mestra"):

        with tempfile.TemporaryDirectory() as pasta_temp:

            for arquivo in arquivos:

                caminho = os.path.join(
                    pasta_temp,
                    arquivo.name
                )

                with open(caminho, "wb") as f:
                    f.write(arquivo.getbuffer())

            excel_gerado = gerar_lista_mestra(
                pasta_temp
            )

            with open(excel_gerado, "rb") as f:

                st.download_button(
                    "Baixar Excel",
                    data=f,
                    file_name=os.path.basename(excel_gerado),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )