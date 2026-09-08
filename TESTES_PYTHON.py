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




pasta_destino = os.chdir(r'C:\Users\marcus.silva05\Desktop\PRODUÇÃO\SED_PREENCHER')
arquivo_excel = 'UDEMO_TESTES.xlsx'
df = pd.read_excel(arquivo_excel, sheet_name=0)  # 0 = primeira aba






