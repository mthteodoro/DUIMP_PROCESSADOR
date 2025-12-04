import streamlit as st
import io
import os
import subprocess
import time

st.set_page_config(page_title="WESTROCK Processador", page_icon="📦")

CAMINHO_BACKEND = r"C:\Users\matheus.teodoro\Desktop\PYTHON\DUIMP_WESTROCK\backend.py"
CAMINHO_ARQUIVO_BACKEND = r"C:\Users\matheus.teodoro\Desktop\PYTHON\Catalogo WESTROCK - DUIMP 12092025.xlsx"
CAMINHO_SAIDA = r"C:\Users\matheus.teodoro\Desktop\PYTHON\Resultado_WESTROCK_PROCESSADO.xlsx"

st.title("📦 Processar Catálogo")
st.write("Envie o arquivo Excel e clique em **Processar Arquivo**.")

uploaded = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"])

if uploaded:

    st.success("Arquivo carregado com sucesso!")

    if st.button("🚀 Processar Arquivo"):

        # 1 - Salvar arquivo no local esperado pelo backend (sem alterar o backend)
        with open(CAMINHO_ARQUIVO_BACKEND, "wb") as f:
            f.write(uploaded.getbuffer())

        st.info("Arquivo salvo. Iniciando processamento...")

        # 2 - Executar o backend como script externo
        try:
            subprocess.run(
                f'python "{CAMINHO_BACKEND}"',
                check=True,
                shell=True,
                capture_output=True,
                text=True
            )
        except subprocess.CalledProcessError as e:
            st.error("Erro ao executar o backend:")
            st.code(e.stdout + "\n" + e.stderr)
            st.stop()

        st.success("Processamento concluído!")

        # 3 - Verificar se o arquivo final existe
        if not os.path.exists(CAMINHO_SAIDA):
            st.error("O backend não gerou o arquivo de saída.")
            st.stop()

        # 4 - Disponibilizar para download
        with open(CAMINHO_SAIDA, "rb") as f:
            st.download_button(
                label="📥 Baixar Resultado",
                data=f,
                file_name="Resultado_WESTROCK_PROCESSADO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


