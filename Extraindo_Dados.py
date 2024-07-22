from selenium import webdriver as opcpesSelenium
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui
import pyautogui as pausa

from openpyxl import load_workbook
import os
import xlsxwriter as opcoesExcel

#Caminho do arquivo no computador (PRECISO SEMPRE CRIAR A TABELA ANTES)
nome_arquivo_tabela = r"C:\Extrações\Mercado_Livre\Dados_ML_Games.xlsx"
planilhaDadosTabela = load_workbook(nome_arquivo_tabela)

#Selecionar a sheet de dados
sheet_dados = planilhaDadosTabela['Games']
print()

#Abre o navegador
nav = opcpesSelenium.Chrome()
nav.get("https://mercadolivre.com.br")
nav.maximize_window()

#Preenche os campos
pausa.sleep(0.5)
nav.find_element(By.NAME, "as_word").send_keys("Jogos ps5")

pausa.sleep(2)
pyautogui.press("enter")
#nav.find_element(By.NAME, "as_word").send_keys(Keys.ENTER)

pausa.sleep(5)

dadosProduto = nav.find_elements(By.CLASS_NAME,"ui-search-result__content-wrapper")

linha = 2
for informacoes in dadosProduto:

    #Selecionar a sheet de dados
    sheet_dados = planilhaDadosTabela['Games']

    nomeProduto = informacoes.find_element(By.CLASS_NAME, "ui-search-item__title").text
    precoProduto = informacoes.find_element(By.CLASS_NAME, "andes-money-amount__fraction").text

    try:
        centavosProduto = informacoes.find_element(By.CLASS_NAME, "andes-money-amount__cents andes-money-amount__cents--superscript-24").text
    except:
        centavosProduto = "0"

    urlProduto = informacoes.find_element(By.TAG_NAME, "a").get_attribute("href")

    #print(nomeProduto + " - " + precoProduto + "," + centavosProduto + " - " + urlProduto)

    #Pegar a última linha + 1
    linha = len(sheet_dados['A']) + 1

    #Demos o nome da coluna + o número da linha
    colunaA = "A" + str(linha) #A2
    colunaB = "B" + str(linha) #B2
    colunaC = "C" + str(linha) #C2

    # Imprimimos os dados na tabela do excel
    sheet_dados["A1"] = "Produto"
    sheet_dados["B1"] = "Preço"
    sheet_dados["C1"] = "Imagem"

    precoTexto = precoProduto + "," + centavosProduto

    precoSemPonto = precoTexto.replace('.', '')
    precoSemPonto2 = precoSemPonto.replace(',', '.')

    #Convertendo para moeda, usando o float.
    precoSemPonto2 = float(precoSemPonto2)

    #Imprimimos os dados na tabela do excel
    sheet_dados[colunaA] = nomeProduto
    sheet_dados[colunaB] = precoSemPonto2
    sheet_dados[colunaC] = urlProduto


#Salva o arquivo com as alterações
planilhaDadosTabela.save(filename=nome_arquivo_tabela)

#Abre o arquivo
os.startfile(nome_arquivo_tabela)


