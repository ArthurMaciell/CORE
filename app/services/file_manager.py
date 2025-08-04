from app.utils.logging_config import logger
import os
import streamlit as st




# Dicionário com os caminhos base por fábrica, vendedor e segmento
CAMINHOS = {
    ("IMI", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\7. NCI (IMI)",
    ("IMI", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\7. NCI (IMI)",

    ("TROX ACESSORIO", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\1. NCT (TROX)",
    ("TROX ACESSORIO", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\1. NCT (TROX)",

    ("TROX EQUIPAMENTO", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\1. NCT-X (TROX Equip.)",
    ("TROX EQUIPAMENTO", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\1. NCT-X (TROX Equip.)",

    ("DAIKIN", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\2. NCD (Daikin)",
    ("DAIKIN", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\2. NCD (Daikin)",

    ("ARMACELL", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\5. NOA (Armacell)",
    ("ARMACELL", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\5. NOA (Armacell)",

    ("PROJELMEC", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\4. NOP (Projelmec)",
    ("PROJELMEC", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\4. NOP (Projelmec)",

    ("ARMSTRONG", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\6. NAM (Armstrong)",
    ("ARMSTRONG", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\6. NAM (Armstrong)",

    ("LEVEROS", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\8. NCS (Outros Equip.)",
    ("LEVEROS", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\8. NCS (Outros Equip.)",

    ("MELTING", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\8. NCS (Outros Equip.)",
    ("MELTING", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\8. NCS (Outros Equip.)",

    ("EVAPCO", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Produtos\3. NCE (Evapco)",
    ("EVAPCO", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Produtos\3. NCE (Evapco)",

    ("SERVIÇO TROX", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Serviços\1. SCT (TROX)",
    ("SERVIÇO TROX", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Serviços\1. SCT (TROX)",

    ("SERVIÇO EVAPCO", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Serviços\3. SCE (Evapco)",
    ("SERVIÇO EVAPCO", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Serviços\3. SCE (Evapco)",

    ("SERVIÇO IMI", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Serviços\7. SCI (IMI)",
    ("SERVIÇO IMI", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Serviços\7. SCI (IMI)",

    ("SERVIÇO ARMSTRONG", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Serviços\6. SAM (Armstrong)",
    ("SERVIÇO ARMSTRONG", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Serviços\6. SAM (Armstrong)",

    ("SERVIÇO PROJELMEC", "Bruno Crispim"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.5 RETEC DF-TO\Serviços\4. SOP (Projelmec)",
    ("SERVIÇO PROJELMEC", "Gabriel Bento"): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\5.6 RETEC GO\Serviços\4. SOP (Projelmec)",

    ("Distribuição", None): r"C:\Users\Orçamento\OneDrive - GRUPO RETEC\02. Engenharia\CTD"
}


def criar_pasta(nome_pasta: str, fabrica_orc: str, vend_orc: str, seguimento_orc: str) -> str:
    """
    Cria a pasta no diretório correto com base na fábrica, vendedor e segmento.
    """

    if not nome_pasta:
        st.error("O nome da pasta não foi gerado corretamente.")
        return ""

    # Verifica o caminho com chave mais específica, senão tenta por seguimento
    caminho_base = CAMINHOS.get((fabrica_orc, vend_orc)) or CAMINHOS.get((seguimento_orc, None))

    if not caminho_base:
        st.error('ERRO! Não foi possível encontrar o caminho base.')
        return ""

    caminho_nova_pasta = os.path.join(caminho_base, nome_pasta)
    st.write(f"Caminho completo: {caminho_nova_pasta}")

    try:
        os.makedirs(caminho_nova_pasta, exist_ok=True)
        st.success(f'Pasta criada com sucesso: {caminho_nova_pasta}')
        return caminho_nova_pasta
    except Exception as e:
        st.error(f"Erro ao criar a pasta: {str(e)}")
        return ""
