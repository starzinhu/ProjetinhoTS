#-*- coding: utf-8 -*-
"""
Automatizador para verificar versões de dispositivos em uma página web.
"""

import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException
from datetime import datetime

def verificar_versoes_dispositivos():
    """
    Acessa a página de relatório de liberados, compara as versões
    e retorna os IDs dos dispositivos com versões iguais na data atual.
    """
    url = "https://172.16.232.219/trocar_versao/relatorio_de_liberados/"
    dispositivos_com_versao_ok = []
    versao_liberada_padrao = "L-02.01.19"
    versao_atual_padrao = "A-02.01.19"
    
    # Obter a data atual no formato esperado (ex: 23/10/2025)
    data_hoje = datetime.now().strftime('%d/%m/%Y')
    
    # Configurações do Chrome para aceitar certificados SSL inválidos
    options = Options()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--headless') # Roda sem abrir a janela do navegador
    options.add_argument('--disable-gpu')

    driver = None # Inicializa a variável
    try:
        driver = webdriver.Chrome(options=options)
        print(f"Acessando a URL: {url}")
        driver.get(url)
        print("Atualizando a página para garantir os dados mais recentes...")
        driver.refresh()
        
        # Aguarda um pouco para a página carregar
        time.sleep(5) 

        print("Procurando pela tabela de dispositivos...")
        # Suposição: A tabela de dados é a primeira encontrada com a tag 'table'
        tabela = driver.find_element(By.TAG_NAME, 'table')
        linhas = tabela.find_elements(By.TAG_NAME, 'tr')

        if len(linhas) <= 1:
            print("Nenhum dado encontrado na tabela.")
            return

        # Suposição: O cabeçalho está na primeira linha e as colunas são fixas
        # Vamos mapear os índices das colunas que nos interessam
        # Isso precisa ser ajustado se a ordem das colunas mudar
        # Dispositivo, Versão Liberada, Versão Atual, Data de Liberação
        # A ordem exata precisa ser confirmada inspecionando a página
        # Por enquanto, vamos supor as seguintes posições:
        # Dispositivo -> coluna 0
        # Versão Liberada -> coluna 1
        # Versão Atual -> coluna 2
        # Data de Liberação -> coluna 3
        
        print(f"Analisando registros para a data de hoje: {data_hoje}")
        for linha in linhas[1:]:
            colunas = linha.find_elements(By.TAG_NAME, 'td')
            # Com base no diagnóstico, agora sabemos que precisamos de 8 colunas
            if len(colunas) >= 8:
                
                # Extrai os textos das colunas com os índices corretos
                dispositivo_id = colunas[0].text.strip()
                versao_liberada = colunas[1].text.strip()
                versao_atual = colunas[2].text.strip()
                # A data de liberação está na 8ª coluna (índice 7)
                data_liberacao_com_hora = colunas[7].text.strip()

                # Pega apenas a parte da data (antes do espaço)
                data_liberacao = data_liberacao_com_hora.split(' ')[0]

                # Compara a data e as versões
                if data_liberacao == data_hoje:
                    # Verifica se a versão liberada e a versão atual correspondem aos padrões definidos
                    if versao_liberada == versao_liberada_padrao and versao_atual == versao_atual_padrao:
                        dispositivos_com_versao_ok.append(dispositivo_id)
                        print(f"  - Dispositivo {dispositivo_id} OK (Liberada: {versao_liberada}, Atual: {versao_atual})")

        if dispositivos_com_versao_ok:
            print("\n--- Dispositivos com a versão atualizada hoje ---")
            print(", ".join(dispositivos_com_versao_ok))
            print(f"\nQuantidade de IDs atualizados: {len(dispositivos_com_versao_ok)}")
        else:
            print(f"\nNenhum dispositivo encontrado com a versão liberada '{versao_liberada_padrao}' e atual '{versao_atual_padrao}' hoje.")

    except WebDriverException as e:
        print(f"Erro ao acessar a página: {e}")
        print("Verifique se a URL está acessível e se o navegador está instalado corretamente.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
    finally:
        if driver:
            driver.quit()
            print("Navegador fechado.")

if __name__ == "__main__":
    verificar_versoes_dispositivos()
