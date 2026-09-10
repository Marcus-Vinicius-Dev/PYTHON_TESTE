# python -m venv venv (cria venv se necessário)
# .\.venv\Scripts\Activate (ativa a venv)
# deactivate (desativa venv)
# pip install -r requirements.txt
# python -m pip install --force-reinstall --no-cache-dir -r requirements.txt
# Start-Process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -ArgumentList "--remote-debugging-port=9222", "--user-data-dir=C:\edge-debug"
# cd C:\Users\vinic\Desktop\Python\WebScraping01
# python TESTES_SCRAPING.py

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



pasta_arquivo = os.chdir(r'C:\Users\marcus.silva05\Desktop\git_projetos\PYTHON_TESTE')
pasta_trecho = os.path.basename(r'C:\Users\marcus.silva05\Desktop\git_projetos\PYTHON_TESTE')
arquivo_excel = 'CHOR_MATTATHIAS.xlsx'
abas = pd.read_excel(arquivo_excel, sheet_name=None)
#print(Fore.GREEN + f"\nUSANDO ARQUIVO '{arquivo_excel}' DA PASTA '{pasta_trecho}'")
#print(Fore.RED + f"\nO ARQUIVO POSSUI {len(abas)} ABAS")
#for i, (nome, df) in enumerate(abas.items(), 1):
#    print(Fore.WHITE + f"  {i}. {nome}")
#    print(Fore.WHITE + f"     Linhas: {len(df)}")
#    print(Fore.WHITE + f"     Colunas: {len(df.columns)}")
#    print(Fore.WHITE + f"     Colunas: {list(df.columns)[:3]}...")
#    print()

aba_selecionada = '2026'  # Nome da aba que deseja selecionar
df = pd.read_excel(arquivo_excel, sheet_name=aba_selecionada)  # Carrega apenas a aba escolhida
# print(Fore.GREEN + f"\n✅ ABA SELECIONADA: '{aba_selecionada}'")
# print(Fore.WHITE + f"   Linhas: {len(df)}")
# print(Fore.WHITE + f"   Total de Colunas: {len(df.columns)}")
# print(Fore.WHITE + f"   Colunas: {list(df.columns)}")


"""
O ARQUIVO POSSUI 1 ABAS
  1. 2026
     Linhas: 128
     Colunas: 10
     Colunas: ['NOMEESC', 'ID_INTERNO', 'CARGO_C']...

✅ ABA SELECIONADA: '2026'
   Linhas: 128
   Total de Colunas: 10
   Colunas: ['NOMEESC', 'ID_INTERNO', 'CARGO_C', 'NM_CARGOC', 'MATERIA', 'DEN_MATERIA', 'TOT_AULA_LIVRE', 'TOT_AULA_SUBST', 'TOT_GERAL_AULA', 'JORNADA']
    
"""

# SELECT * FROM df
print(Fore.GREEN + "\n📊 SELECT *:")
print(Fore.WHITE + f"   Total de registros: {len(df)}")
print(Fore.WHITE + f"   Total de colunas: {len(df.columns)}")
print(Fore.CYAN + f"\n{df}")

###############################################################################################


substituicoes = {'CAO': 'ÇÃO', 'ME': 'MÉ', 'BAS': 'BÁS', 'II':'2', 'PROFESSOR':'PROFESSOR DE'}  # Dicionário de substituições
df['NOVA_COLUNA'] = df['NM_CARGOC'].replace(substituicoes, regex=True)  # Substitui todos de uma vez
print(Fore.GREEN + f"\n📊 SUBSTITUIÇÕES MÚLTIPLAS:")
print(Fore.WHITE + f"   Dicionário: {substituicoes}")
print(Fore.CYAN + f"{df[['NM_CARGOC', 'NOVA_COLUNA']].head()}")






