# Web Scraping Portal SSP

# Autor: Rafael Freitas
# Data: 01/06/2025

# Bibliotecas necessárias

# pip install pandas
# pip install selenium
# pip install webdriver-manager
# pip install tkinter

# Contesto do projeto: 

# Coletar dados do site da SSP-SP (Secretaria de Segurança Pública do Estado de São Paulo)
# e tratar esses dados para análise posterior.
# O site em questão é: https://www.ssp.sp.gov.br/estatistica/dados-mensais
# O objetivo é extrair dados de criminalidade por bairro na cidade de São Paulo
# para os anos especificados pelo usuário.
# Os dados serão salvos em um arquivo Excel para análise posterior.
# O código também inclui tratamento de exceções para lidar com possíveis erros durante a execução.
# O código utiliza Selenium para automação do navegador e Pandas para manipulação de dados.
# O código é interativo, solicitando ao usuário o ano inicial e final para a coleta de dados.
# O código filtra os dados para incluir apenas a cidade de São Paulo e bairros específicos.
# O código salva os dados coletados em um arquivo Excel na pasta do usuário.
# O código é projetado para ser executado em um ambiente local com acesso à internet e ao navegador Chrome.
# O código utiliza o gerenciador de drivers do Chrome para garantir que a versão correta do driver seja usada.

import os
import time
import pandas as pd
import tkinter as tk
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Criar janela para entrada do usuário
root = tk.Tk()
root.withdraw()

# Perguntar ao usuário qual ano inicial e final deseja pesquisar
ano_inicial = simpledialog.askinteger("Coleta de Dados", "Digite o ano inicial da pesquisa:")
ano_final = simpledialog.askinteger("Coleta de Dados", "Digite o ano final da pesquisa:")

if ano_inicial and ano_final:
    print(f"Pesquisando dados do ano {ano_inicial} até {ano_final}")

    # WebDriver.
    options = Options()
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        url = "https://www.ssp.sp.gov.br/estatistica/dados-mensais"
        driver.get(url)

        # O site tem um sistema de loading próprio, por isso coloquei
        # um time.sleep(3) para que o código não trave.

        time.sleep(3)

        # lista de armazenar dos dados.
        dados_completos = []

        # execução do código com base na respostas das perguntas.
        for ano in range(ano_inicial, ano_final + 1):
            print(f"Selecionando o ano: {ano}")

            # Seleciona o ano.
            seletor_ano = Select(WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/app-root/body/div[1]/div/app-dados-mensais/div[2]/div[2]/form/div[3]/div[2]/select"))
            ))
            seletor_ano.select_by_visible_text(str(ano))

            # Sempre irá selecionar a região "Capital".
            # No futuro é possível criar um novo filtro de pergunta.
            seletor_regiao = Select(WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/app-root/body/div[1]/div/app-dados-mensais/div[2]/div[2]/form/div[3]/div[3]/select"))
            ))
            seletor_regiao.select_by_visible_text("Capital")

            # Sempre irá selecionar cidade "São Paulo".
            # OBS.: É necessário fazer essa seleção para que site disponibilize
            # a seleção de delegacias.

            seletor_cidade = Select(WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/app-root/body/div[1]/div/app-dados-mensais/div[2]/div[2]/form/div[3]/div[4]/select"))
            ))
            seletor_cidade.select_by_visible_text("São Paulo")

            # Irá selecionar a delegacia, más para facilitar o entendimento
            # dos dados eu troquei o nome "Delegacias" por "Bairro".
            seletor_bairro = Select(WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/app-root/body/div[1]/div/app-dados-mensais/div[2]/div[2]/form/div[3]/div[5]/select"))
            ))
            lista_bairros = [bairro.text.strip() for bairro in seletor_bairro.options if bairro.text.strip()]

            # Irá percorrer todos os bairros.
            for bairro in lista_bairros:
                try:
                    seletor_bairro.select_by_visible_text(bairro)
                    print(f"Consultando dados para o bairro: {bairro}")
                    time.sleep(1)

                    # Coleta e define as colunas.
                    cabecalho = driver.find_elements(By.XPATH, "/html/body/app-root/body/div[1]/div/app-dados-mensais/div[2]/div[2]/form/div[4]/div/div[1]/div/div/div/table/thead/tr/th")
                    titulos = ["Ano", "Bairro"] + [titulo.text.strip() for titulo in cabecalho if titulo.text.strip()]

                    # Coleta os dados.
                    linhas = driver.find_elements(By.XPATH, "/html/body/app-root/body/div[1]/div/app-dados-mensais/div[2]/div[2]/form/div[4]/div/div[1]/div/div/div/table/tbody/tr")

                    # Devido a uma particularidade do site, tive que
                    # coletar os dados dessa forma.
                    for linha in linhas:
                        # Captura os valores da primeira coluna (Natureza).
                        natureza = linha.find_element(By.TAG_NAME, "th").text.strip()

                        # Captura os valores das demais colunas.
                        colunas = linha.find_elements(By.TAG_NAME, "td")
                        dados_linha = [coluna.text.strip() for coluna in colunas]

                        # Garante que todas as linhas tenham o mesmo número de colunas.
                        while len(dados_linha) < len(titulos) - 3:  # Ajustando para incluir "Ano", "Bairro" e "Natureza".
                            dados_linha.append("")

                        # Adiciona as colunas "Ano" e "Bairro".
                        dados_completos.append([ano, bairro, natureza] + dados_linha)

                except Exception:
                    print(f"Aviso: O bairro '{bairro}' não pôde ser selecionado. Pulando...")
                    continue

        # Cria um DataFrame e salva como Excel direto na pasta do Drive
        df = pd.DataFrame(dados_completos, columns=titulos)
        nome_arquivo = "Dados_Portal_SSP.xlsx"
        df.to_excel(f"{nome_arquivo}", index=False)

        print(f"Os dados da tabela foram exportados com sucesso para {nome_arquivo}!")

    except Exception as e:
        print("Ocorreu um erro ao tentar realizar a operação:", e)

    driver.quit()

else:
    print("Ano inicial ou final inválido.")