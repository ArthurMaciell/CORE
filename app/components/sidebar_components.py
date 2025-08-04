import streamlit as st
import os
import base64
import openpyxl
from app.services.spreadsheet_manager import processar_plan


VENDEDORES_CTD = ["Iago Rangel", "Luan Araujo", "Marlon Souza", "Rutemar Junior", "Wellisson Chaves"]


def sidebar_distribuicao(caminho_pasta, vend_orc, nome_orc, numero_orc, revisao_orc):
    st.sidebar.subheader('Planilha Distribuição')
    
    plans_tqs = st.sidebar.file_uploader('Faça upload do EXCEL do TQS aqui:', accept_multiple_files=False)

    if plans_tqs:
        wb1 = openpyxl.load_workbook(plans_tqs)
        ws1 = wb1.worksheets[1]
        # Opcional: st.sidebar.success('Planilha carregada!')

    if st.button('Processar Planilha'):
        with st.spinner('Processando...'):
            novo_arquivo = os.path.join(os.getcwd(), 'PlanDist.xlsx')
            processar_plan(plans_tqs, nome_orc, vend_orc, numero_orc, revisao_orc, novo_arquivo)

            with open(novo_arquivo, 'rb') as f:
                b64 = base64.b64encode(f.read()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="CTD_{numero_orc}_REV_{revisao_orc}.xlsx">Download Planilha Cotação</a>'
                st.markdown(href, unsafe_allow_html=True)


def sidebar_upload_arquivos(caminho_pasta):
    st.sidebar.subheader('Upload de Arquivos')
    uploaded_files = st.sidebar.file_uploader('Faça upload dos arquivos:', accept_multiple_files=True)

    if uploaded_files and st.sidebar.button('Salvar Arquivos'):
        for file in uploaded_files:
            with open(os.path.join(caminho_pasta, file.name), 'wb') as f:
                f.write(file.getbuffer())
        st.sidebar.success("Arquivos salvos com sucesso!")

    if uploaded_files:
        nomes = [f.name for f in uploaded_files]
        arquivo = st.sidebar.selectbox('Qual arquivo quer renomear?', nomes)
        novo_nome = st.sidebar.text_input('Novo nome do arquivo') + '.pdf'

        if st.sidebar.button('Renomear Arquivo'):
            caminho_antigo = os.path.join(caminho_pasta, arquivo)
            caminho_novo = os.path.join(caminho_pasta, novo_nome)
            try:
                os.rename(caminho_antigo, caminho_novo)
                st.sidebar.success(f"{arquivo} renomeado para {novo_nome}")
            except Exception as e:
                st.sidebar.error(f"Erro ao renomear: {e}")
