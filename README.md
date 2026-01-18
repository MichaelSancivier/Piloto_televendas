# 🚛 Michelin Pilot Command Center

Ferramenta de processamento de dados para balanceamento de carteiras, higienização de telefones e geração de kits de discagem.

## 🚀 Funcionalidades Principais
- **Balanceamento Automático**: Distribui clientes órfãos e nivela a carga de trabalho entre atendentes.
- **Higienização de Dados**: Limpa IDs e formata telefones para o padrão internacional (+55).
- **Segmentação por Perfil**: Separa automaticamente Frotistas (Manhã) e Freteiros (Almoço).
- **Verticalização de Discador**: Transforma colunas de telefone em registros individuais para aumentar a contactabilidade.
- **Reforço de Tarde**: Filtra clientes contactados via Log do Discador e gera lista de remanescentes.

## 📋 Pré-requisitos
- Python 3.8+
- Bibliotecas: `streamlit`, `pandas`, `openpyxl`, `xlsxwriter`

## 🛠️ Como Instalar
1. Clone este repositório.
2. Instale as dependências:
   ```bash
   pip install streamlit pandas openpyxl xlsxwriter
