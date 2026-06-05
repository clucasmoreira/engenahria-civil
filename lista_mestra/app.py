import streamlit as st
import tempfile
import os

from lista_mestra import gerar_lista_mestra

st.set_page_config(
    page_title="Lista Mestra",
    layout="wide"
)

st.title("📑 Gerador de Lista Mestra")

st.write(
    "Selecione todos os PDFs e TXTs do projeto."
)

arquivos = st.file_uploader(
    "Arquivos",
    type=["pdf", "txt"],
    accept_multiple_files=True
)

if arquivos:

    st.subheader("Arquivos recebidos")

    for arquivo in arquivos:
        st.write(f"✓ {arquivo.name}")

    qtd_pdf = sum(
        1 for arq in arquivos
        if arq.name.lower().endswith(".pdf")
    )

    qtd_txt = sum(
        1 for arq in arquivos
        if arq.name.lower().endswith(".txt")
    )

    st.info(
        f"{qtd_pdf} PDF(s) e {qtd_txt} TXT(s)"
    )

    if st.button("Gerar Lista Mestra"):

        try:

            with tempfile.TemporaryDirectory() as pasta_temp:

                # salva todos os arquivos enviados

                for arquivo in arquivos:

                    caminho = os.path.join(
                        pasta_temp,
                        arquivo.name
                    )

                    with open(caminho, "wb") as f:
                        f.write(
                            arquivo.getbuffer()
                        )

                # DEBUG
                st.write("Arquivos gravados:")

                for nome in os.listdir(
                    pasta_temp
                ):
                    st.write(nome)

                # gera excel

                excel_gerado = gerar_lista_mestra(
                    pasta_temp
                )

                st.success(
                    "Lista mestra criada com sucesso!"
                )

                with open(
                    excel_gerado,
                    "rb"
                ) as f:

                    st.download_button(
                        label="📥 Baixar Excel",
                        data=f,
                        file_name=os.path.basename(
                            excel_gerado
                        ),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as erro:

            st.error(
                f"Erro ao gerar lista: {erro}"
            )
