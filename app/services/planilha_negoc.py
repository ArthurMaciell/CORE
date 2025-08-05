import pandas as pd
import streamlit as st



def plan_neg(cliente):
    caminho_plan = r"C:\Users\Orçamento\ONE DRIVE ORCAMENTO\OneDrive - GRUPO RETEC\02. Engenharia\Dep. Orçamentos\POWERBI\AUTOMACAO RD\data\negociacoes_2025.xlsx"
    df = pd.read_excel(caminho_plan)
    
    # Verificando colunas disponíveis
    colunas_desejadas = [
        'Proposta N°',
        'deal_stage.name',
        'Fator',
        'name',
        'amount_total',
        'closed_at',
        'organization.name',
        'user.name',
        'Nome da Obra',
        'Produtos (Distribuição)',
        'Produtos (Representação)',
        'Orçamentista'
    ]

    # Filtrando apenas as colunas desejadas (verificando se elas existem)
    colunas_existentes = [col for col in colunas_desejadas if col in df.columns]
    df = df[colunas_existentes]    
    
    if cliente:
        # Filtro por cliente (você pode ajustar para ficar mais robusto, como .str.contains())
        df_filtrado = df[df['organization.name'].str.contains(cliente, case=False, na=False)]

        st.subheader(f"Negociações do Cliente: {cliente}")
        if not df_filtrado.empty:
            st.dataframe(df_filtrado)
        else:
            st.warning("Nenhuma negociação encontrada para esse cliente.")
            
        # Contagem dos estágios da negociação
        st.subheader("Contagem por Estágio de Negociação")
        contagem_estagios = df_filtrado['deal_stage.name'].value_counts().reset_index()
        contagem_estagios.columns = ['Estágio da Negociação', 'Quantidade']
        st.dataframe(contagem_estagios)
        # Substitui vírgula por ponto
        df_filtrado['Fator'] = df_filtrado['Fator'].astype(str).str.replace(',', '.', regex=False)
        df_filtrado['Fator'] = pd.to_numeric(df_filtrado['Fator'], errors='coerce')
        media_fator = (
            df_filtrado
            .groupby('deal_stage.name', dropna=True)['Fator']
            .mean()
            .reset_index()
            .sort_values(by='Fator', ascending=False)
        )
        media_fator.columns = ['Estágio da Negociação', 'Média do Fator']
        st.dataframe(media_fator)
        