
from app.utils.logging_config import logger



def nomear_pasta(vend_orc, fabrica_orc, numero_orc, cliente_orc, obra_orc, data_orc_conc, revisao_orc, seguimento_orc,nome_orc):
    if fabrica_orc == 'TROX EQUIPAMENTO':
        codigo_orc = 'NCT-X'
    elif fabrica_orc == 'TROX ACESSORIO':
        codigo_orc = 'NCT'
    elif fabrica_orc == 'DAIKIN':
        codigo_orc = 'NCD'
    elif fabrica_orc == 'ARMACELL':
        codigo_orc = 'NOA'
    elif fabrica_orc == 'IMI':
        codigo_orc = 'NCI'
    elif fabrica_orc == 'PROJELMEC':
        codigo_orc = 'NOP'
    elif fabrica_orc == 'LEVEROS':
        codigo_orc = 'NCS'
    elif fabrica_orc == 'ARMSTRONG':
        codigo_orc = 'NAM'
    elif fabrica_orc == 'EVAPCO':
        codigo_orc = 'NCE'
    elif fabrica_orc == 'MELTING':
        codigo_orc = 'NCO'
    elif fabrica_orc == 'SERVIÇO TROX':
        codigo_orc = 'SCT'
    elif fabrica_orc == 'SERVIÇO EVAPCO':
        codigo_orc = 'SCE'
    elif fabrica_orc == 'SERVIÇO IMI':
        codigo_orc = 'SCI'
    elif fabrica_orc == 'SERVIÇO ARMSTRONG':
        codigo_orc = 'SAM' 
    elif fabrica_orc == 'SERVIÇO PROJELMEC':
        codigo_orc = 'SOP'      


        
    else:
        codigo_orc = 'NAO IDENTIFIQUEI O CÓDIGO'   
        
    if vend_orc == 'Gabriel Bento':
        vend_orc = 'GB'
        
    elif vend_orc == 'Bruno Crispim':
        vend_orc = 'BC'
        
    elif vend_orc == 'Iago Rangel':
        vend_orc = 'IR'
        
    elif vend_orc == 'Luan Araujo':
        vend_orc = 'LA'
        
    elif vend_orc == 'Marlon Souza':
        vend_orc = 'MA'
        
    elif vend_orc == 'Rutemar Junior':
        vend_orc = 'RJ'
        
    elif vend_orc == 'Wellisson Chaves':
        vend_orc = 'WC'
    
    elif vend_orc == 'Larissa Sousa':
        vend_orc = 'LS'


    #Nomeando os arquivos da Distribuição
    if seguimento_orc == 'Distribuição':
        codigo_orc = 'CTD'
    
    
    data_formatada = data_orc_conc.replace('/', '.')
    nome_pasta = codigo_orc + ' ' + numero_orc + '-25 - ' + cliente_orc + ' - ' + obra_orc + ' - ' + data_formatada + ' - ' + vend_orc
    nome_arquivo = codigo_orc + ' ' + numero_orc + '-25 - ' + data_formatada + ' - ' + vend_orc
    if revisao_orc != '0':
        nome_pasta = codigo_orc + ' ' + numero_orc  + '-25' +'- REV'+ revisao_orc + ' - ' + cliente_orc + ' - ' + obra_orc + ' - ' + data_formatada + ' - ' + vend_orc
    
    logger.info(f"Gerando nome para orçamento: N°={numero_orc}, fábrica={fabrica_orc}, cliente={cliente_orc},orçamentista={nome_orc}")
    #st.write(nome_pasta)
    #st.write(fabrica_orc)
    return(nome_pasta, nome_arquivo)