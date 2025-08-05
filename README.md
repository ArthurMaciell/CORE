# CORE – Cadastro de Orçamentos com Registro e Extração Automatizada
> Sistema para automação de orçamentos comerciais. Gera planilhas automaticamente, organiza pastas e arquivos com base em parâmetros como vendedor e localização, e cadastra os orçamentos em banco de dados para posterior análise e integração com ferramentas como Power BI.

![Python](https://img.shields.io/badge/Python-3.10-blue)
![Streamlit](https://img.shields.io/badge/Framework-Streamlit-ff4b4b)
![Status](https://img.shields.io/badge/status-Em%20desenvolvimento-yellow)


---

## 🚀 Sobre o Projeto

O **CORE (Cadastro de Orçamentos com Registro Eficiente)** é um sistema interno construído com Python e Streamlit, que automatiza tarefas repetitivas e críticas do setor de orçamentos.

Com ele, é possível:

- Criar pastas automaticamente com nomes estruturados
- Realizar upload de planilhas do TQS
- Interpretar códigos de itens por meio de **Regex**
- Gerar planilhas de venda com base nas extrações
- Registrar data, hora e responsável pelo orçamento
- Calcular o tempo útil gasto na elaboração
- Salvar os dados automaticamente em planilhas e diretórios organizados

---

## 🧠 Tecnologias Utilizadas

- **Python 3.10**
- **Streamlit** – interface web
- **Regex (re)** – para extração inteligente de códigos
- **Pandas / Openpyxl** – para manipulação de planilhas Excel
- **Datetime** – para cálculo de tempo útil
- **OS / Pathlib** – para manipulação de arquivos e diretórios
- **OneDrive** – para armazenamento compartilhado

---

## 💻 Como Executar Localmente

1. Clone o repositório:
```bash
git clone https://github.com/ArthurMaciell/CORE.git
cd core-orcamentos
