import streamlit as st
import locale
import datetime
from app.services.nomeador import nomear_pasta
from app.utils.tempo_util import calcular_horas_uteis,calcular_horas_uteis_simples
from app.services.file_manager import criar_pasta
from app.utils.logging_config import logger




#Nome da Página
st.set_page_config(page_title='CORE 1.0',page_icon=':open_file_folder:', layout="centered")

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
vend_orc = st.selectbox("Qual é o Vendedor?",["Gabriel Bento","Bruno Crispim", "Iago Rangel","Luan Araujo","Marlon Souza","Rutemar Junior", "Wellisson Chaves"])
data_orc_ini = st.date_input('Data de Inicio (D/M/A)')
hora_orc_ini = st.text_input('Horário de Inicio (HH:MM)')
#data_orc_conc = st.text_input('Data de Conclusão(D/M/A)')
revisao_orc = st.text_input('Qual é o número/letra da revisão?')
if vend_orc in ["Gabriel Bento","Bruno Crispim"]:
    fator_orc = st.text_input('Qual é o Fator do orçamento?')
elif vend_orc in ["Iago Rangel","Luan Araujo","Marlon Souza","Rutemar Junior", "Wellisson Chaves"]:
    fator_orc = '1'
valor_orc = st.number_input('Qual é o Valor do Orçamento? (DIGITE O PREÇO, NÃO COLOQUE PONTO, USE VÍRGULA) ')

#Organização para ver se vai ser por Representação ou Distribuição
if vend_orc in ["Iago Rangel","Luan Araujo","Marlon Souza","Rutemar Junior", "Wellisson Chaves"]:
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
    st.success(f"Nome da pasta: {nome_pasta}")
else:
    st.error('Erro ao nomear a pasta.')
    


try:
    horas_uteis = calcular_horas_uteis(data_orc_ini, hora_orc_ini, data_orc_conc, hora_atual_plan)
    st.success(f"Horas úteis gastas: {horas_uteis}")
    logger.info(f"Orçamento finalizado: {nome_pasta} | Horas úteis: {horas_uteis}")
except ValueError:
    st.error("Data ou hora em formato inválido. Verifique os campos e tente novamente.")
    logger.info(f"Erro ao calcular as horas úteis.")


# No seu botão de Streamlit:
if st.button("Criar Pasta"):
    caminho_final = criar_pasta(nome_pasta, fabrica_orc, vend_orc, seguimento_orc)


