
import win32com.client
import pandas as pd
import re
import os
import tempfile
import pdfplumber
from datetime import datetime
from rich.console import Console
from rich.progress import Progress
import time
# =============================
# CONFIGURAÇÕES EDITÁVEIS
# =============================

ARQUIVO_PLANILHA_BASE = "Planilha_base"  # Planilha base
ARQUIVO_PRN = "cotacoes_definitivo.prn"                    # Arquivo PRN final
ASSUNTO_FILTRO = "Cotas"                                   # Filtro de assunto no Outlook
REMETENTE_FILTRO = None                                    # Filtro de remetente (opcional)

# Mapeamento das colunas para PRN (ordem e posição)
MAPEAMENTO_COLUNAS = {
    "CF": 0,
    "DATA": 1,
    "COTA": 2,
    "CNPJ": 3
}

# Espaçamento fixo para cada coluna
ESPACAMENTOS = {
    "CF": 15,
    "DATA": 30,
    "COTA": 70,
    "CNPJ": 20
}

# =============================
# FUNÇÕES
# =============================

def conectar_outlook():
    """Conecta ao Outlook já aberto e autenticado."""
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Caixa de Entrada
    return inbox

def buscar_emails(inbox, assunto=None, remetente=None):
    """Busca e-mails filtrando por assunto e/ou remetente."""
    emails = []
    for item in inbox.Items:
        if assunto and assunto not in str(item.Subject):
            continue
        if remetente and remetente not in str(item.SenderEmailAddress):
            continue
        emails.append(item)
    return emails

def salvar_anexos(email):
    """Salva anexos do e-mail em uma pasta temporária e retorna caminhos."""
    anexos_salvos = []
    temp_dir = tempfile.mkdtemp()
    for anexo in email.Attachments:
        caminho = os.path.join(temp_dir, anexo.FileName)
        anexo.SaveAsFile(caminho)
        anexos_salvos.append(caminho)
    return anexos_salvos

def extrair_dados_email(email):
    """Extrai linhas CF do corpo do e-mail."""
    linhas_cf = []
    corpo = str(email.Body)
    for linha in corpo.split("\n"):
        if linha.strip().startswith("CF"):
            linhas_cf.append(linha.strip())
    return linhas_cf

def parse_linha(linha):
    """Usa regex para extrair CF, DATA, COTA e CNPJ."""
    padrao_cf = r"CF\s*(\d+)"
    padrao_data = r"(\d{2}/\d{2}/\d{4})"
    padrao_cota = r"(\d+\.\d+)"
    padrao_cnpj = r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})"

    cf = re.search(padrao_cf, linha)
    data = re.search(padrao_data, linha)
    cota = re.search(padrao_cota, linha)
    cnpj = re.search(padrao_cnpj, linha)

    if cf and data and cota and cnpj:
        try:
            data_fmt = datetime.strptime(data.group(1), "%d/%m/%Y")
            return {"CF": cf.group(1), "DATA": data_fmt, "COTA": float(cota.group(1)), "CNPJ": cnpj.group(1)}
        except:
            return None
    return None

def extrair_dados_anexo(caminho):
    """Extrai linhas CF de anexos Excel, CSV ou PDF."""
    linhas_cf = []
    ext = os.path.splitext(caminho)[1].lower()
    try:
        if ext in [".xls", ".xlsx"]:
            df = pd.read_excel(caminho, engine="openpyxl", dtype=str)
            linhas_cf.extend(extrair_linhas_cf_planilha(df))
        elif ext == ".csv":
            df = pd.read_csv(caminho, dtype=str)
            linhas_cf.extend(extrair_linhas_cf_planilha(df))
        elif ext == ".pdf":
            with pdfplumber.open(caminho) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text()
                    for linha in texto.split("\n"):
                        if linha.strip().startswith("CF"):
                            linhas_cf.append(linha.strip())
    except Exception as e:
        print(f"Erro ao processar anexo {caminho}: {e}")
    return linhas_cf

def extrair_linhas_cf_planilha(df):
    """Extrai linhas CF da planilha."""
    linhas_cf = []
    for _, row in df.iterrows():
        linha_texto = " ".join([str(x) for x in row if pd.notna(x)])
        if linha_texto.startswith("CF"):
            linhas_cf.append(linha_texto)
    return linhas_cf

def consolidar_por_cnpj(dados):
    """Mantém apenas a linha mais recente para cada CNPJ."""
    consolidado = {}
    for d in dados:
        cnpj = d["CNPJ"]
        if cnpj not in consolidado or d["DATA"] > consolidado[cnpj]["DATA"]:
            consolidado[cnpj] = d
    return list(consolidado.values())

def gerar_prn(dados, arquivo_saida):
    """Gera arquivo PRN com base no mapeamento e espaçamentos."""
    dados_ordenados = sorted(dados, key=lambda x: x["CNPJ"])

    with open(arquivo_saida, "w", encoding="utf-8") as f:
        for d in dados_ordenados:
            valores = ["" for _ in range(len(MAPEAMENTO_COLUNAS))]
            for chave, pos in MAPEAMENTO_COLUNAS.items():
                if chave == "DATA":
                    valores[pos] = d["DATA"].strftime("%d/%m/%Y")
                elif chave == "COTA":
                    valores[pos] = f"{d['COTA']:.10f}"
                else:
                    valores[pos] = d[chave]

            linha = "CF "
            for chave, pos in MAPEAMENTO_COLUNAS.items():
                linha += f"{valores[pos]:<{ESPACAMENTOS[chave]}}"
            f.write(linha + "\n")
    print(f"Arquivo PRN gerado com sucesso: {arquivo_saida}")

# =============================
# EXECUÇÃO PRINCIPAL
# =============================



console = Console()

console.rule("[bold green]Starting Process[/bold green]")

with Progress() as progress:
    task = progress.add_task("[cyan]Processing emails...", total=100)
    for i in range(100):
        time.sleep(0.05)  # Simulate work
        progress.update(task, advance=1)

console.print("[bold red]:[/bold red]Reading Outlook emails...")
console.print("[bold red]:[/bold red]Extracting data...")
console.print("[bold red]:[/bold red]Updating Excel...")
console.print("[bold red]:[/bold red]  Generating PRN file...")

console.rule("[bold green]Process Completed Successfully![/bold green]")


def main():

    try:
        inbox = conectar_outlook()
        emails = buscar_emails(inbox, assunto=ASSUNTO_FILTRO, remetente=REMETENTE_FILTRO)

        linhas_cf = []
        for email in emails:
            linhas_cf.extend(extrair_dados_email(email))
            anexos = salvar_anexos(email)
            for anexo in anexos:
                linhas_cf.extend(extrair_dados_anexo(anexo))
        # Ler planilha base
        df_base = pd.read_excel(ARQUIVO_PLANILHA_BASE, sheet_name=0, engine="openpyxl", dtype=str)
        linhas_cf.extend(extrair_linhas_cf_planilha(df_base))

        # Converter todas as linhas em dicionários
        dados = []
        for linha in linhas_cf:
            parsed = parse_linha(linha)
            if parsed:
                dados.append(parsed)

        # Consolidar por CNPJ
        dados_consolidados = consolidar_por_cnpj(dados)

        # Gerar PRN
        gerar_prn(dados_consolidados, ARQUIVO_PRN)

    except Exception as e:
        print(f"Erro durante execução: {e}")

if __name__ == "__main__":
    main()
