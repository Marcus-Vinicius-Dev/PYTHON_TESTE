# deactivate (desativa venv)
# !!!SE NECESSÁRIO!!! Remove-Item -Recurse -Force venv (Remover o .venv)
# python -m venv venv (cria venv)
# venv\Scripts\Activate.ps1 (ativa a venv)
# pip install -r requirements.txt
# python -m pip install --force-reinstall --no-cache-dir -r requirements.txt
# Start-Process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -ArgumentList "--remote-debugging-port=9222", "--user-data-dir=C:\edge-debug"
# cd C:\Users\vinic\Desktop\GIT\PYTHON_TESTE
# cd C:\Users\marcus.silva05\Desktop\git_projetos\PYTHON_TESTE
# python TESTES_PYTHON.py

# sys
from colorama import init, Fore, Back, Style
from datetime import datetime
import datetime
import os
import platform
import pkgutil
import psutil
import requests
import shutil
import socket
import subprocess
import sys
import wmi

# leitura
from docx import Document
from openpyxl import load_workbook
import csv
import json
import pandas as pd
import re
import unicodedata
import tarfile 
import zipfile   

# debug
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pyautogui
import pygetwindow as gw
import uiautomation as auto
import time

init(autoreset=True) # reseta a cor no próximo print


###################################################################### CHOR_MATTATHIAS ######################################################################
pasta_arquivo = os.chdir(r'C:\Users\marcus.silva05\Desktop\git_projetos\PYTHON_TESTE') 
pasta_trecho = os.path.basename(f'{pasta_arquivo}') 
arquivo_excel = 'CHOR_MATTATHIAS.xlsx'
abas = pd.read_excel(arquivo_excel, sheet_name=None)
# print(Fore.GREEN + f"\nUSANDO ARQUIVO '{arquivo_excel}' DA PASTA '{pasta_trecho}'")
# print(Fore.RED + f"\nO ARQUIVO POSSUI {len(abas)} ABA(S)")
# for i, (nome, df) in enumerate(abas.items(), 1):
#     print(Fore.WHITE + f"  {i}. {nome}")
#     print(Fore.WHITE + f"     Linhas: {len(df)}")
#     print(Fore.WHITE + f"     Colunas: {len(df.columns)}")
#     print(Fore.WHITE + f"     Colunas: {list(df.columns)[:3]}...")
#     print()

aba_selecionada = 'CHOR_2026'  # Nome da aba que deseja selecionar
df = pd.read_excel(arquivo_excel, sheet_name=aba_selecionada)  # Carrega apenas a aba escolhida
# print(Fore.GREEN + f"\n✅ ABA SELECIONADA: '{aba_selecionada}'")
# print(Fore.WHITE + f"   Linhas: {len(df)}")
# print(Fore.WHITE + f"   Total de Colunas: {len(df.columns)}")
# print(Fore.WHITE + f"   Colunas: {list(df.columns)}")


# SELECT * FROM df
# print(Fore.GREEN + "\n📊 SELECT *:")
# print(Fore.WHITE + f"   Total de registros: {len(df)}")
# print(Fore.WHITE + f"   Total de colunas: {len(df.columns)}")
# print(Fore.CYAN + f"\n{df}")


"""
USANDO ARQUIVO 'CHOR_MATTATHIAS.xlsx' DA PASTA 'PYTHON_TESTE'

O ARQUIVO POSSUI 1 ABA(S)
  1. 2026
     Linhas: 128
     Colunas: 10
     Colunas: ['NOMEESC', 'ID_INTERNO', 'CARGO_C']...


✅ ABA SELECIONADA: '2026'
   Linhas: 128
   Total de Colunas: 10
   Colunas: ['NOMEESC', 'ID_INTERNO', 'CARGO_C', 'NM_CARGOC', 'MATERIA', 'DEN_MATERIA', 'TOT_AULA_LIVRE', 'TOT_AULA_SUBST', 'TOT_GERAL_AULA', 'JORNADA']

📊 SELECT *:
   Total de registros: 128
   Total de colunas: 10
    
"""
###################################################################### CHOR_MATTATHIAS ######################################################################

###################################################################### CTG_MATTATHIAS ######################################################################
pasta_arquivo_2 = os.chdir(r'C:\Users\marcus.silva05\Desktop\git_projetos\PYTHON_TESTE') 
pasta_trecho_2 = os.path.basename(f'{pasta_arquivo}') 
arquivo_excel_2 = 'CTG_MATTATHIAS.xlsx'
abas_2 = pd.read_excel(arquivo_excel_2, sheet_name=None)
# print(Fore.GREEN + f"\nUSANDO ARQUIVO '{arquivo_excel_2}' DA PASTA '{pasta_trecho_2}'")
# print(Fore.RED + f"\nO ARQUIVO POSSUI {len(abas_2)} ABA(S)")
# for i, (nome, df_2) in enumerate(abas_2.items(), 1):
#     print(Fore.WHITE + f"  {i}. {nome}")
#     print(Fore.WHITE + f"     Linhas: {len(df_2)}")
#     print(Fore.WHITE + f"     Colunas: {len(df_2.columns)}")
#     print(Fore.WHITE + f"     Colunas: {list(df_2.columns)[:3]}...")
#     print()

aba_selecionada_2 = 'MATTATHIAS_CLASSIFICADOS'  # Nome da aba que deseja selecionar
df_2 = pd.read_excel(arquivo_excel_2, sheet_name=aba_selecionada_2)  # Carrega apenas a aba escolhida
# print(Fore.GREEN + f"\n✅ ABA SELECIONADA: '{aba_selecionada_2}'")
# print(Fore.WHITE + f"   Linhas: {len(df_2)}")
# print(Fore.WHITE + f"   Total de Colunas: {len(df_2.columns)}")
# print(Fore.WHITE + f"   Colunas: {list(df_2.columns)}")

aba_selecionada_3 = 'MATTATHIAS_EXERCICIO'  # Nome da aba que deseja selecionar
df_3 = pd.read_excel(arquivo_excel_2, sheet_name=aba_selecionada_3)  # Carrega apenas a aba escolhida
# print(Fore.GREEN + f"\n✅ ABA SELECIONADA: '{aba_selecionada_3}'")
# print(Fore.WHITE + f"   Linhas: {len(df_3)}")
# print(Fore.WHITE + f"   Total de Colunas: {len(df_3.columns)}")
# print(Fore.WHITE + f"   Colunas: {list(df_3.columns)}")

"""
###################################################################### MATTATHIAS_CLASSIFICADOS ######################################################################

  2. MATTATHIAS_CLASSIFICADOS
     Linhas: 85
     Colunas: 36
     Colunas: ['POLO_C', 'REGIAO_C', 'NOMEDE_C']...

  ✅ ABA SELECIONADA: 'MATTATHIAS_CLASSIFICADOS'
   Linhas: 85
   Total de Colunas: 36
   Colunas: ['POLO_C', 'REGIAO_C', 'NOMEDE_C', 'UAC', 'NOMEUA_C', 'MUNICIPIO_C', 'DI', 'SEXO', 'DATA_NASCIMENTO', 'IDADE', 'CARGO_C', 'NOMECAR_C', 'QUADRO_C', 'CATEG_C', 'DT_INICIO_EXERCICIO_C',
    'ANOS_TRAB_CARGO_C', 'DTPOSSE_C', 'DISC_CONCURSO', 'POLO_E', 'REGIAO_E', 'NOMEDE_E', 'UA_E', 'LOCAL_UA', 'NOMEUA_E', 'MUNICIPIO_E', 'CARGO_E', 'NOMECAR_E', 'QUADRO_E', 'CATEG_E', 
    'DATA_INICIO_EXERCICIO_E', 'DTPOSSE_E', 'QTDE_ATS', 'JORNADA', 'TIPO_DEF', 'ID_COR', 'ID_INTERNO']  


  3. MATTATHIAS_EXERCICIO
     Linhas: 83
     Colunas: 36
     Colunas: ['POLO_C', 'REGIAO_C', 'NOMEDE_C']...

  ✅ ABA SELECIONADA: 'MATTATHIAS_EXERCICIO'
   Linhas: 83
   Total de Colunas: 36
   Colunas: ['POLO_C', 'REGIAO_C', 'NOMEDE_C', 'UAC', 'NOMEUA_C', 'MUNICIPIO_C', 'DI', 'SEXO', 'DATA_NASCIMENTO', 'IDADE', 'CARGO_C', 'NOMECAR_C', 'QUADRO_C', 'CATEG_C', 'DT_INICIO_EXERCICIO_C',
    'ANOS_TRAB_CARGO_C', 'DTPOSSE_C', 'DISC_CONCURSO', 'POLO_E', 'REGIAO_E', 'NOMEDE_E', 'UA_E', 'LOCAL_UA', 'NOMEUA_E', 'MUNICIPIO_E', 'CARGO_E', 'NOMECAR_E', 'QUADRO_E', 'CATEG_E', 
    'DATA_INICIO_EXERCICIO_E', 'DTPOSSE_E', 'QTDE_ATS', 'JORNADA', 'TIPO_DEF', 'ID_COR', 'ID_INTERNO']    
  
###################################################################### MATTATHIAS_CLASSIFICADOS ######################################################################

"""

###################################################################### CTG_MATTATHIAS ######################################################################


print(Fore.GREEN + f"Lista de Colunas de {aba_selecionada}")
for i, col in enumerate(df.columns, 1):
    print(Fore.CYAN + f"   {i:3d}. {col}")
print(Fore.WHITE + f"   Total: {len(df.columns)} colunas")

print(Fore.GREEN + f"Lista de Colunas de {aba_selecionada_2}")
for i, col in enumerate(df_2.columns, 1):
    print(Fore.CYAN + f"   {i:3d}. {col}")
print(Fore.WHITE + f"   Total: {len(df_2.columns)} colunas")

print(Fore.GREEN + f"Lista de Colunas de {aba_selecionada_3}")
for i, col in enumerate(df_3.columns, 1):
    print(Fore.CYAN + f"   {i:3d}. {col}")
print(Fore.WHITE + f"   Total: {len(df_3.columns)} colunas")



df_final = pd.merge(df, df_2, on='ID_INTERNO', how='left', suffixes=('_A', '_B')) \
    .assign( # how='left'ou 'right' ou 'inner' ou 'outter' (outer = full)
        CARGO=lambda x: x['CARGO_C_A'].astype(str) + ' - ' + x['NOMECAR_C'] # CONCATENAÇÃO
    ) \
    .rename(columns={
        'REGIAO_C': 'REGIAO',           # sem sufixo (só no df_2)
        'NOMEDE_C': 'URE',              # sem sufixo (só no df_2)
        'NOMEESC': 'ESCOLA',            # sem sufixo (só no df)
        'ID_INTERNO': 'ID',           # ← sufixo (existe nos dois A e B)
        'CATEG_C': 'CATEGORIA',         # sem sufixo (só no df_2)
        'DISC_CONCURSO': 'DISCIPLINA',  # sem sufixo (só no df_2)
        'DEN_MATERIA': 'MATERIA',       # sem sufixo (só no df)
        'JORNADA_A': 'JORNADA',         # ← sufixo _A (existe nos dois)
        'TOT_GERAL_AULA': 'TOTAL_AULAS' # sem sufixo (só no df)
    }) \
    [['REGIAO', 'URE', 'ESCOLA', 'ID', 'CARGO', 'CATEGORIA',
      'DISCIPLINA', 'MATERIA', 'JORNADA', 'TOTAL_AULAS']] # SEQUÊNCIA 

print(Fore.GREEN + f"\n📊 DATAFRAME FINAL:")
print(Fore.WHITE + f"   Registros: {len(df_final)}")
print(Fore.CYAN + f"\n{df_final.head(10)}")