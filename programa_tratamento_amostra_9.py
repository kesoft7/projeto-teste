import pandas as pd
import re
from pathlib import Path
import tkinter as tk
from tkinter import simpledialog
import openpyxl
import xlsxwriter

# 1) Janela para entrada do prefixo contábil
root = tk.Tk()
root.withdraw()
prefixo_contabil = simpledialog.askstring("Filtro Contábil", "Digite os primeiros números do código contábil (ex: 1.1.1):")
if not prefixo_contabil:
    raise ValueError("Nenhum prefixo contábil foi informado.")

# 2) Detectar arquivo Excel na pasta
arquivos_excel = list(Path(".").glob("*.xls*"))
if len(arquivos_excel) == 0:
    raise FileNotFoundError("Nenhum arquivo Excel encontrado na pasta do script.")
elif len(arquivos_excel) > 1:
    raise RuntimeError("Mais de um arquivo Excel encontrado. Deixe apenas um na pasta.")
else:
    arquivo = arquivos_excel[0]

print(f"📂 Usando arquivo: {arquivo.name}")

# 3) Carregar o arquivo
df = pd.read_excel(arquivo)

# 4) Limpeza inicial
df = df.iloc[6:].reset_index(drop=True)
df = df.dropna(how="all").reset_index(drop=True)

# 5) Remover cabeçalhos internos, preservando o primeiro
cabecalho_padrao = ["Data", "Partida", "Complemento", "Doc.", "C.Custo", "Débitos", "Créditos"]

def linha_eh_cabecalho(row):
    valores = [str(x).strip() for x in row.tolist()]
    return valores[:len(cabecalho_padrao)] == cabecalho_padrao

# Detectar o índice do primeiro cabeçalho
indice_primeiro_cabecalho = None
for i, row in df.iterrows():
    if linha_eh_cabecalho(row):
        indice_primeiro_cabecalho = i
        break

# Remover todos os cabeçalhos exceto o primeiro
df = df.drop([
    i for i, row in df.iterrows()
    if linha_eh_cabecalho(row) and i != indice_primeiro_cabecalho
]).reset_index(drop=True)

# 6) Criar colunas extras
nova_coluna_codigo = []
nova_coluna_nome = []

codigo_atual = None
nome_atual = None

padrao = re.compile(r"^\d\.\d\.\d\.\d{2}\.\d{2}\.\d{3}$")
linhas_para_remover = set()

for i, row in df.iterrows():
    valor_coluna_b = str(row.iloc[1]).strip()
    valor_coluna_d = str(row.iloc[3]).strip()

    if padrao.match(valor_coluna_b):
        codigo_atual = valor_coluna_b
        nome_atual = valor_coluna_d
        linhas_para_remover.add(i)

    nova_coluna_codigo.append(codigo_atual)
    nova_coluna_nome.append(nome_atual)

df["Novo Código"] = pd.Series(nova_coluna_codigo).fillna("")
df["Novo Nome"] = pd.Series(nova_coluna_nome).fillna("")

# 7) Remover linhas com código/nome original
df = df.drop(index=linhas_para_remover).reset_index(drop=True)

# 8) Reordenar colunas
colunas = ["Novo Código", "Novo Nome"] + [c for c in df.columns if c not in ["Novo Código", "Novo Nome"]]
df = df[colunas]

# 9) Filtrar lançamentos com prefixo contábil
df_lancamentos = df[df["Novo Código"].fillna("").str.startswith(prefixo_contabil)].reset_index(drop=True)

# 10) Criar DataFrame com cabeçalho personalizado
cabecalho_personalizado = ["Conta", "Nome da Conta", "Data", "ID_Lanc", "Histórico", "Código", "C. Custo", "Débito", "Crédito"]
df_lancamentos.columns = cabecalho_personalizado[:len(df_lancamentos.columns)]

# 11) Salvar em novo arquivo com duas abas
saida = arquivo.with_name(arquivo.stem + " - Transformado.xlsx")
with pd.ExcelWriter(saida, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Transformado", index=False)
    df_lancamentos.to_excel(writer, sheet_name="Lançamentos", index=False)

print(f"✅ Planilha transformada salva como: {saida}")