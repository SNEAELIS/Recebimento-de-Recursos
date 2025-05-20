import logging
import os

from openpyxl.reader.excel import load_workbook
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
from webdriver_manager.chrome import ChromeDriverManager
from xlsxwriter import Workbook


def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        logging.info("Conectado ao navegador existente!")
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        logging.error(f"Erro ao conectar ao navegador: {erro}")
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()


def ler_planilha_entrada(caminho_arquivo=r"C:\Users\diego.brito\Downloads\robov1\pasta1.xlsx"):
    try:
        dataframe = pd.read_excel(caminho_arquivo, engine="openpyxl")

        if "Instrumento nº" in dataframe.columns:
            dataframe["Instrumento nº"] = (
                dataframe["Instrumento nº"]
                .astype(str)
                .str.replace(r"\.0$", "", regex=True)
            )
            return dataframe
    except Exception as erro:
        print(f"❌ Erro ao ler planilha de entrada: {erro}")
    exit()



def esperar_elemento_por_xpath(navegador, xpath, tempo_limite=10, modo='clicavel'):
    try:
        if modo == 'clicavel':
            return WebDriverWait(navegador, tempo_limite).until(
                EC.element_to_be_clickable((By.XPATH, xpath)))
        elif modo == 'visivel':
            return WebDriverWait(navegador, tempo_limite).until(
                EC.visibility_of_element_located((By.XPATH, xpath)))
        else:  # default: apenas presente no DOM
            return WebDriverWait(navegador, tempo_limite).until(
                EC.presence_of_element_located((By.XPATH, xpath)))
    except Exception as erro:
        print(f"❌ Elemento não encontrado ou interativo: {xpath} - {str(erro)}")
        return None


def esperar_elemento_xpath(navegador, xpath, tempo=10):
    try:
        return WebDriverWait(navegador, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except Exception as e:
        print(f"❌ Elemento com XPath '{xpath}' não encontrado: {e}")
        return None


def esperar_elemento_css(navegador, seletor, tempo=10):
    try:
        return WebDriverWait(navegador, tempo).until(EC.element_to_be_clickable((By.CSS_SELECTOR, seletor)))
    except Exception as e:
        print(f"❌ Elemento com seletor CSS '{seletor}' não encontrado com ou não clicável: {e}")
        return None




def navegar_para_instrumento(navegador, numero_instrumento):
    """Navega até a página do instrumento específico"""
    try:
        esperar_elemento_por_xpath(navegador, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]").click()
        esperar_elemento_por_xpath(navegador,
                                   "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]").click()

        campo_pesquisa = esperar_elemento_por_xpath(
            navegador,
            "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]"
        )
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(numero_instrumento)

        esperar_elemento_por_xpath(
            navegador,
            "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]"
        ).click()

        time.sleep(1)
        esperar_elemento_por_xpath(
            navegador,
            "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]"
        ).click()

        return True
    except Exception:
        print(f"⚠️ Instrumento {numero_instrumento} não encontrado.")
        return False



def verificar_e_registrar_repasses(navegador, instrumento_id):
    base_path = "base_dados.xlsx"
    saida_path = "resultados_atuais.xlsx"

    # Clique no primeiro menu
    esperar_elemento_css(navegador, "#div_-173460853 > span > span").click()
    time.sleep(1)

    # Clique no segundo menu
    esperar_elemento_css(navegador, "#menu_link_-173460853_-503637981 > div > span > span").click()
    time.sleep(1)

    # Tentar clicar no botão de acessar dados de pagamento
    botao_detalhe = esperar_elemento_css(navegador, "#tbodyrow > tr > td:nth-child(6) > nobr > a")
    if not botao_detalhe:
        registrar_excel(saida_path, {
            "Instrumento": instrumento_id,
            "Status": "Dados de pagamento não encontrados"
        })
        return

    botao_detalhe.click()
    time.sleep(2)

    # Esperar o conteúdo da aba de repasses carregar
    esperar_elemento_xpath(navegador, '//*[@id="ConteudoDiv"]')

    # Coletar os valores
    try:
        valor_previsto = navegador.find_element(By.ID, "tr-inserirOBConfluxoValorPrevisto").text.split("R$")[-1].strip()
        valor_desembolsado = navegador.find_element(By.ID, "tr-inserirOBConfluxoValorDesembolsado").text.split("R$")[-1].strip()
        valor_a_desembolsar = navegador.find_element(By.ID, "tr-inserirOBConfluxoValorADesembolsar").text.split("R$")[-1].strip()
        situacao = navegador.find_element(By.XPATH, '//*[@id="tbodyrow"]/tr/td[9]/div').text.strip()
        data_emissao = navegador.find_element(By.XPATH, '//*[@id="tbodyrow"]/tr/td[10]/div').text.strip()
    except Exception as e:
        print(f"⚠️ Erro ao extrair valores: {e}")
        return

    dados_novos = {
        "Instrumento": instrumento_id,
        "Valor Previsto": valor_previsto,
        "Valor Desembolsado": valor_desembolsado,
        "Valor a Desembolsar": valor_a_desembolsar,
        "Situação": situacao,
        "Data de Emissão da OB": data_emissao,
        "Status": "Coletado"
    }

    # Carrega base anterior, compara e salva se for novo ou alterado
    df_novos = pd.DataFrame([dados_novos])
    if os.path.exists(base_path):
        df_base = pd.read_excel(base_path, engine="openpyxl")

        filtro = df_base["Instrumento"] == instrumento_id
        if not filtro.any():
            df_base = pd.concat([df_base, df_novos], ignore_index=True)
        else:
            dados_antigos = df_base.loc[filtro].iloc[0].to_dict()
            for chave in ["Valor Previsto", "Valor Desembolsado", "Valor a Desembolsar", "Situação", "Data de Emissão da OB"]:
                if dados_antigos.get(chave) != dados_novos.get(chave):
                    print(f"📌 Alteração detectada no instrumento {instrumento_id}: {chave}")
                    df_base.loc[filtro, chave] = dados_novos[chave]

        df_base.to_excel(base_path, index=False)
    else:
        df_novos.to_excel(base_path, index=False)

    # Atualizar arquivo de saída da execução atual
    registrar_excel(saida_path, dados_novos)


def registrar_excel(caminho_arquivo, dados):
    if not os.path.exists(caminho_arquivo):
        wb = Workbook()
        ws = wb.active
        ws.append(list(dados.keys()))
        wb.save(caminho_arquivo)

    wb = load_workbook(caminho_arquivo)
    ws = wb.active
    ws.append(list(dados.values()))
    wb.save(caminho_arquivo)





def main():
    print("🚀 Iniciando o robô Selenium...")

    # Conectar ao navegador existente
    navegador = conectar_navegador_existente()

    # Ler a planilha de entrada
    df_entrada = ler_planilha_entrada()

    # Verificar se há dados na coluna "Instrumento nº"
    if df_entrada.empty or "Instrumento nº" not in df_entrada.columns:
        print("❌ A planilha não contém dados válidos na coluna 'Instrumento nº'. Encerrando...")
        return

    # Iterar sobre os instrumentos e navegar
    for idx, linha in df_entrada.iterrows():
        numero_instrumento = linha["Instrumento nº"]
        print(f"🔎 Buscando instrumento: {numero_instrumento}")

        sucesso = navegar_para_instrumento(navegador, numero_instrumento)

        if sucesso:
            print(f"✅ Instrumento {numero_instrumento} acessado com sucesso!")
            verificar_e_registrar_repasses(navegador, numero_instrumento)
        else:
            print(f"⚠️ Não foi possível acessar o instrumento {numero_instrumento}.")
            registrar_excel("resultados_atuais.xlsx", {
                "Instrumento": numero_instrumento,
                "Status": "Instrumento não encontrado"
            })

        time.sleep(2)  # espera entre execuções

    print("🏁 Finalizado!")




