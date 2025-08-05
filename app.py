import streamlit as st
import locale
from datetime import datetime, timedelta
from app.services.nomeador import nomear_pasta
from app.utils.tempo_util import calcular_horas_uteis,calcular_horas_uteis_simples
from app.services.file_manager import criar_pasta
from app.utils.logging_config import logger
from app.components.sidebar_components import (sidebar_distribuicao, sidebar_upload_arquivos,VENDEDORES_CTD)
from app.services.orcamento_cadastro import cadastrar_orcamento
from datetime import datetime


#Nome da Página
st.set_page_config(page_title='CORE 1.0',page_icon=':open_file_folder:', layout="centered")

# Organização do Menu Lateral (Sidebar)
st.sidebar.image('imagens\LOGO RETEC-Photoroom.png', width=200)
st.sidebar.header("Gerenciar Arquivos")

# Configurar o locale para português (Brasil)
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

#Input Usuário
nome_orc = st.selectbox("Qual é o nome do Orçamentista?",["Arthur Maciel","Felipe Almeida","Luiz Henrique","Matheus Fernandes"])
numero_orc = st.text_input('Número de Orçamento') 
terceiros_orc = st.checkbox('O orçamento depende de Terceiros?')
cliente_orc = st.text_input('Nome do Cliente') 
icms_orc = st.selectbox('O Cliente é Contribuinte ICMS?',['Não Contribuinte','Contribuinte','Isento'])
obra_orc = st.text_input('Nome da Obra')
local_obra = st.text_input('Local da Obra (CIDADE/UF)') 
frete_orc = st.selectbox('Qual é o Tipo de frete?',['FOB','CIF'])
tipo_orc = st.selectbox('Qual é o Tipo de Orçamento? ',['Concorrência', 'Compra','Consulta','Pedido de Compra']) 
vend_orc = st.selectbox("Qual é o Vendedor?",["Gabriel Bento","Bruno Crispim", "Iago Rangel","Luan Araujo","Larissa Sousa","Rutemar Junior", "Wellisson Chaves"])
data_orc_ini = st.date_input('Data de Inicio (D/M/A)')
hora_orc_ini = st.text_input('Horário de Inicio (HH:MM)')
#data_orc_conc = st.text_input('Data de Conclusão(D/M/A)')
revisao_orc = st.text_input('Qual é o número/letra da revisão?')
if vend_orc in ["Gabriel Bento","Bruno Crispim"]:
    fator_orc = st.text_input('Qual é o Fator do orçamento?')
elif vend_orc in ["Iago Rangel","Luan Araujo","Larissa Sousa","Rutemar Junior", "Wellisson Chaves"]:
    fator_orc = '1'
valor_orc = st.number_input('Qual é o Valor do Orçamento? (DIGITE O PREÇO, NÃO COLOQUE PONTO, USE VÍRGULA) ')

#Organização para ver se vai ser por Representação ou Distribuição
if vend_orc in ["Iago Rangel","Luan Araujo","Larissa Sousa","Rutemar Junior", "Wellisson Chaves"]:
    fabrica_orc = st.selectbox("Qual é a Fábrica Distribuição?",["TROX ACESSORIO","TROX EQUIPAMENTO", "IMI","ARMACELL","PENNSE","DAIKIN", "SCIFLUX", 'OUTROS'])
    seguimento_orc = "Distribuição"

elif vend_orc in ["Gabriel Bento","Bruno Crispim"]:
    fabrica_orc = st.selectbox("Qual é a Fábrica Representação?",["TROX ACESSORIO","TROX EQUIPAMENTO", "IMI","ARMACELL","PROJELMEC","ARMSTRONG","DAIKIN", "EVAPCO", 'LEVEROS', 'MELTING', 'SERVIÇO','OUTROS'])
    seguimento_orc = "Representação"

if fabrica_orc == 'SERVIÇO':
    fabrica_orc = st.selectbox('Qual é a Fabrica do Serviço?', ['SERVIÇO TROX','SERVIÇO EVAPCO','SERVIÇO IMI','SERVIÇO ARMSTRONG','SERVIÇO PROJELMEC'])

tamanho_orc = st.selectbox("Qual é o Tamanho do orçamento?",["Pequeno","Médio", "Grande", 'Gigante'])

# Conversões de data
data_orc_ini = data_orc_ini.strftime('%d/%m/%Y')
hora_atual_plan = datetime.now().strftime('%H:%M')
data_orc_conc = datetime.today().strftime("%d/%m/%Y")

# Geração dos nomes
if st.button("Gerar nome da pasta"):
    nome_pasta, nome_arquivo = nomear_pasta(
        vend_orc, fabrica_orc, numero_orc, cliente_orc, obra_orc, data_orc_conc, revisao_orc, seguimento_orc
    )
    st.session_state['nome_pasta'] = nome_pasta
    st.session_state['nome_arquivo'] = nome_arquivo
    st.success(f"Nome da pasta: {nome_pasta}")
#else:
    #st.error('Erro ao nomear a pasta.')
    


    try:
        horas_uteis = calcular_horas_uteis(data_orc_ini, hora_orc_ini, data_orc_conc, hora_atual_plan)
        st.success(f"Horas úteis gastas: {horas_uteis}")
        logger.info(f"Orçamento finalizado: {nome_pasta} | Horas úteis: {horas_uteis}")
    except ValueError:
        st.error("Data ou hora em formato inválido. Verifique os campos e tente novamente.")
        logger.info(f"Erro ao calcular as horas úteis.")

# Checagem do estado e botão para criar pasta
if st.button('Criar Pasta'):
    if not st.session_state.get('nome_pasta'):
        st.warning("Antes de criar a pasta, gere o nome primeiro.")
    else:
        nome_pasta = st.session_state['nome_pasta']
        
        # Criação da pasta
        caminho_pasta = criar_pasta(
            nome_pasta=nome_pasta,
            fabrica_orc=fabrica_orc,
            vend_orc=vend_orc,
            seguimento_orc=seguimento_orc
        )
        
        if caminho_pasta:
            st.session_state['caminho_pasta'] = caminho_pasta
            st.success("Pasta criada com sucesso.")
            st.write(caminho_pasta)
        else:
            st.error("Falha ao criar a pasta.")

# Se a pasta já foi criada e está no session_state, exibe o menu de upload ou distribuição
if st.session_state.get('caminho_pasta'):
    caminho_pasta = st.session_state['caminho_pasta']

    if vend_orc in VENDEDORES_CTD:
        sidebar_distribuicao(caminho_pasta, vend_orc, nome_orc, numero_orc, revisao_orc)
    else:
        sidebar_upload_arquivos(caminho_pasta)

dias_semana = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"]

if st.button('Cadastrar Orçamento'):
    try:
        path_planilha = r"C:\Users\Orçamento\ONE DRIVE ORCAMENTO\OneDrive - GRUPO RETEC\02. Engenharia\Dep. Orçamentos\CADASTRO ORÇAMENTO RETEC\CORE\data\Cadastro Orçamento PYTHON.xlsx"
        
        date_datetime_ini = datetime.strptime(data_orc_ini, "%d/%m/%Y")
        date_datetime_conc = datetime.strptime(data_orc_conc, "%d/%m/%Y")

        dados = {
            "data_orc_ini": data_orc_ini,
            "hora_orc_ini": hora_orc_ini,
            "data_orc_conc": data_orc_conc,
            "mes_por_extenso_ini": date_datetime_ini.strftime('%B'),
            "dia_semana_ini": dias_semana[date_datetime_ini.weekday()],
            "mes_por_extenso": date_datetime_conc.strftime('%B'),
            "dia_semana": dias_semana[date_datetime_conc.weekday()],
            "nome_orc": nome_orc,
            "fabrica_orc": fabrica_orc,
            "terceiros_orc": terceiros_orc,
            "frete_orc": frete_orc,
            "tipo_orc": tipo_orc,
            "loja_orc": 'GO' if vend_orc in ['Gabriel Bento','Iago Rangel', 'Rutemar Junior'] else 'DF',
            "vend_orc": vend_orc,
            "seguimento_orc": seguimento_orc,
            "cliente_orc": cliente_orc,
            "icms_orc": icms_orc,
            "obra_orc": obra_orc,
            "local_obra": local_obra,
            "tamanho_orc": tamanho_orc,
            "numero_orc": numero_orc,
            "revisao_orc": revisao_orc,
            "fator_orc": fator_orc,
            "valor_orc": valor_orc
        }

        sucesso, resultado = cadastrar_orcamento(path_planilha, dados)

        if sucesso:
            st.success("Cadastro realizado com sucesso!")
            st.write(f"Esse é o orçamento número: {resultado}")
            st.balloons()
        else:
            st.error(f"Erro ao cadastrar: {resultado}")

    except Exception as e:
        st.error(f"Erro inesperado: {str(e)}")


