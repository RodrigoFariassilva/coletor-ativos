
# 📊 Coletor de Ativos Financeiros

Automação para coleta, consolidação e exportação de dados financeiros a partir de múltiplas fontes (Outlook, Excel, CSV, PDF), com geração de arquivo PRN para integração com sistemas legados.

---

## 🚀 Funcionalidades
- Conexão com **Outlook** para leitura de e-mails e anexos.
- Processamento de arquivos **Excel, CSV e PDF**.
- Extração de dados via **Regex** (CF, DATA, COTA, CNPJ).
- Consolidação por **CNPJ** (mantendo a data mais recente).
- Geração de arquivo **PRN** com espaçamento fixo.
- Feedback visual com **Rich** (barra de progresso e mensagens coloridas).

---

## 🛠 Tecnologias
- Python 3.x
- Pandas
- pdfplumber
- pywin32
- openpyxl
- Rich

---

## 📦 Instalação
Clone o repositório:
```bash
git clone https://github.com/RodrigoFariassilva/coletor-ativos.git
cd coletor-ativos
