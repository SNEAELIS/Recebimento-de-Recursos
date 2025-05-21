import logging
import os
import shutil
from openpyxl import load_workbook, Workbook
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
from webdriver_manager.chrome import ChromeDriverManager

# Configuração de logging
logging.basicConfig(level=logging.INFO)

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
        else:
            print("❌ Coluna 'Instrumento nº' não encontrada na planilha.")
            exit()
    except Exception as erro:
        print(f"❌ Erro ao ler planilha de entrada: {erro}")
        exit()

def formatar_data(data_bruta):
    """Tenta formatar a data para o padrão dia/mês/ano"""
    try:
        return pd.to_datetime(data_bruta, dayfirst=True).strftime('%d/%m/%Y')
    except:
        return data_bruta

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
        print(f"❌ Elemento com seletor CSS '{seletor}' não encontrado ou não clicável: {e}")
        return None

def navegar_para_instrumento(navegador, numero_instrumento):
    """Navega até a página do instrumento específico ou retorna ao menu principal"""
    try:
        # Navegar para o menu principal (se já não estiver lá)
        navegador.get("https://discricionarias.transferegov.sistema.gov.br/voluntarias/Principal/Principal.do")
        print(f"🔙 Retornando ao menu principal para o instrumento {numero_instrumento}")

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
    saida_path = r"C:\Users\diego.brito\Downloads\robov1\Recursos\Resultado_Recursos.xlsx"

    # Clique no primeiro menu
    print("  🔍 Clicando no primeiro menu...")
    primeiro_menu = esperar_elemento_css(navegador, "#div_-173460853 > span > span")
    if not primeiro_menu:
        print("  ⚠️ Primeiro menu não encontrado.")
        registrar_excel(saida_path, [{
            "Instrumento": instrumento_id,
            "Status": "Primeiro menu não encontrado"
        }])
        return []
    primeiro_menu.click()
    time.sleep(1)

    # Clique no segundo menu
    print("  🔍 Clicando no segundo menu...")
    segundo_menu = esperar_elemento_css(navegador, "#menu_link_-173460853_-503637981 > div > span > span")
    if not segundo_menu:
        print("  ⚠️ Segundo menu não encontrado.")
        registrar_excel(saida_path, [{
            "Instrumento": instrumento_id,
            "Status": "Segundo menu não encontrado"
        }])
        return []
    segundo_menu.click()
    time.sleep(1)

    # Clicar no botão de detalhes
    print("  🔍 Clicando no botão de detalhes...")
    botao_detalhe = esperar_elemento_css(navegador, "#tbodyrow > tr > td:nth-child(6) > nobr > a")
    if not botao_detalhe:
        print("  ⚠️ Botão de detalhes não encontrado.")
        registrar_excel(saida_path, [{
            "Instrumento": instrumento_id,
            "Status": "Dados de pagamento não encontrados"
        }])
        return []
    botao_detalhe.click()
    time.sleep(2)

    # Verificar se o conteúdo foi carregado
    print("  🔍 Verificando se o conteúdo foi carregado...")
    conteudo = esperar_elemento_xpath(navegador, '//*[@id="ConteudoDiv"]')
    if not conteudo:
        print("  ⚠️ Conteúdo não carregado.")
        registrar_excel(saida_path, [{
            "Instrumento": instrumento_id,
            "Status": "Conteúdo não carregado"
        }])
        return []

    # Valores totais
    try:
        print("  🔍 Extraindo valores totais...")
        valor_previsto = navegador.find_element(By.ID, "tr-inserirOBConfluxoValorPrevisto").text.split("R$")[-1].strip()
        valor_desembolsado = navegador.find_element(By.ID, "tr-inserirOBConfluxoValorDesembolsado").text.split("R$")[-1].strip()
        valor_a_desembolsar = navegador.find_element(By.ID, "tr-inserirOBConfluxoValorADesembolsar").text.split("R$")[-1].strip()
        print(f"  ✅ Valores totais extraídos: Previsto={valor_previsto}, Desembolsado={valor_desembolsado}, A Desembolsar={valor_a_desembolsar}")
    except Exception as e:
        print(f"⚠️ Erro ao extrair valores totais: {e}")
        registrar_excel(saida_path, [{
            "Instrumento": instrumento_id,
            "Status": f"Erro ao extrair valores totais: {str(e)}"
        }])
        return []

    # Lista de linhas da tabela de repasses
    try:
        print("  🔍 Extraindo tabela de repasses...")
        linhas_repasses = navegador.find_elements(By.XPATH, '//*[@id="tbodyrow"]/tr')
        dados_lista = []

        for linha in linhas_repasses:
            try:
                celulas = linha.find_elements(By.TAG_NAME, "td")
                if len(celulas) >= 10:  # Verifica se há colunas suficientes
                    numero_ob = celulas[3].text.strip()  # Número da OB
                    valor = celulas[6].text.strip().replace("R$", "").strip()  # Valor
                    situacao = celulas[8].text.strip()  # Situação
                    data_emissao = celulas[9].text.strip()  # Data de Emissão da OB

                    dados = {
                        "Instrumento": instrumento_id,
                        "Valor Previsto": valor_previsto,
                        "Valor Desembolsado": valor_desembolsado,
                        "Valor a Desembolsar": valor_a_desembolsar,
                        "Número da OB": numero_ob,
                        "Valor Repassado": valor,
                        "Situação": situacao,
                        "Data de Emissão da OB": formatar_data(data_emissao),
                        "Status": "Coletado"
                    }
                    dados_lista.append(dados)
                    print(f"  ✅ Linha extraída: {dados}")
                else:
                    print("  ⚠️ Linha da tabela com colunas insuficientes.")
            except Exception as e:
                print(f"⚠️ Erro ao processar linha da tabela: {e}")
                continue

        if not dados_lista:
            print("  ⚠️ Nenhuma linha de repasse coletada.")
            registrar_excel(saida_path, [{
                "Instrumento": instrumento_id,
                "Status": "Nenhuma linha de repasse coletada"
            }])
        return dados_lista

    except Exception as e:
        print(f"❌ Erro ao localizar linhas da tabela: {e}")
        registrar_excel(saida_path, [{
            "Instrumento": instrumento_id,
            "Status": f"Erro ao localizar tabela: {str(e)}"
        }])
        return []

def registrar_excel(caminho_arquivo, dados_lista):
    print(f"📝 Função registrar_excel chamada com {len(dados_lista)} registros.")
    if not dados_lista:
        print("⚠️ Nada a salvar no Excel (lista vazia).")
        return

    try:
        os.makedirs(os.path.dirname(caminho_arquivo), exist_ok=True)
        if not os.path.exists(caminho_arquivo):
            wb = Workbook()
            ws = wb.active
            ws.title = "Financeiro"
            ws.append([
                "Instrumento nº", "Tipo de Dado", "Valor", "Data de Emissão da OB",
                "Número da OB", "Situação", "Valor Previsto", "Valor Desembolsado",
                "Valor a Desembolsar"
            ])
            wb.save(caminho_arquivo)
            print("ℹ️ Arquivo criado com a aba 'Financeiro'")

        # Converter dados_lista para DataFrame
        df = pd.DataFrame(dados_lista)

        # Ajustar colunas para combinar com a estrutura da aba Financeiro
        df = df.rename(columns={
            "Instrumento": "Instrumento nº",
            "Valor Repassado": "Valor",
            "Status": "Tipo de Dado"
        })

        # Se "Tipo de Dado" não estiver presente, preencher com "Coletado"
        if "Tipo de Dado" not in df.columns:
            df["Tipo de Dado"] = "Coletado"

        # Carregar o arquivo existente
        try:
            df_existente = pd.read_excel(caminho_arquivo, sheet_name="Financeiro", engine="openpyxl")
        except Exception:
            df_existente = pd.DataFrame(columns=[
                "Instrumento nº", "Tipo de Dado", "Valor", "Data de Emissão da OB",
                "Número da OB", "Situação", "Valor Previsto", "Valor Desembolsado",
                "Valor a Desembolsar"
            ])

        # Combinar dados existentes com novos
        df_completo = pd.concat([df_existente, df], ignore_index=True)

        # Salvar os dados na aba Financeiro
        with pd.ExcelWriter(caminho_arquivo, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df_completo.to_excel(writer, sheet_name="Financeiro", index=False)

        print(f"✅ Dados salvos em {caminho_arquivo}")

    except Exception as e:
        print(f"❌ Erro ao salvar no Excel: {e}")
        # Tentar salvar um backup
        backup_path = caminho_arquivo.replace(".xlsx", "_backup.xlsx")
        try:
            pd.DataFrame(dados_lista).to_excel(backup_path, index=False)
            print(f"⚠️ Backup salvo em: {backup_path}")
        except Exception as backup_error:
            print(f"❌ Erro ao salvar backup: {backup_error}")

def converter_valor_monetario(valor):
    """Converte um valor monetário formatado (ex.: '700.000,00') para float."""
    try:
        # Remover 'R$' e espaços, substituir '.' por '' e ',' por '.'
        valor = str(valor).replace("R$", "").strip()
        valor = valor.replace(".", "").replace(",", ".")
        return float(valor)
    except Exception as e:
        print(f"⚠️ Erro ao converter valor monetário '{valor}': {e}")
        return None

def comparar_resultados(df_novo, caminho_base_antiga):
    import pandas as pd
    import os

    if isinstance(caminho_base_antiga, pd.DataFrame):
        df_antigo = caminho_base_antiga
    elif isinstance(caminho_base_antiga, str):
        if not os.path.exists(caminho_base_antiga):
            print("\n📄 [COMPARAÇÃO] Nenhum arquivo anterior encontrado. Primeira coleta.\n")
            return
        df_antigo = pd.read_excel(caminho_base_antiga, sheet_name="Financeiro", engine="openpyxl")
    else:
        print("❌ Tipo de arquivo_antigo inválido.")
        return

    print("\n🔍 [COMPARAÇÃO] Verificando diferenças com o banco anterior...")

    # Verificar colunas necessárias
    colunas_necessarias = ["Instrumento nº", "Valor", "Data de Emissão da OB"]
    colunas_comparacao = ["Valor Previsto", "Valor Desembolsado", "Valor a Desembolsar", "Situação"]
    for col in colunas_necessarias:
        if col not in df_novo.columns or col not in df_antigo.columns:
            print(f"❌ Coluna '{col}' não encontrada em um dos DataFrames. Comparação cancelada.")
            return

    # Padronizar tipos
    df_novo = df_novo.copy()
    df_antigo = df_antigo.copy()

    df_novo["Instrumento nº"] = df_novo["Instrumento nº"].astype(str)
    df_antigo["Instrumento nº"] = df_antigo["Instrumento nº"].astype(str)

    # Padronizar a coluna de data
    df_novo["Data de Emissão da OB"] = df_novo["Data de Emissão da OB"].apply(formatar_data)
    df_antigo["Data de Emissão da OB"] = df_antigo["Data de Emissão da OB"].apply(formatar_data)

    # Converter valores monetários para float
    for col in ["Valor", "Valor Previsto", "Valor Desembolsado", "Valor a Desembolsar"]:
        if col in df_novo.columns and col in df_antigo.columns:
            df_novo[col] = df_novo[col].apply(converter_valor_monetario)
            df_antigo[col] = df_antigo[col].apply(converter_valor_monetario)

    chave = ["Instrumento nº", "Valor", "Data de Emissão da OB"]

    # Encontrar registros novos
    df_merged = df_novo.merge(df_antigo, on=chave, how="left", indicator=True)
    df_novos = df_merged[df_merged["_merge"] == "left_only"]

    if not df_novos.empty:
        print(f"\n🆕 [NOVOS REGISTROS] {len(df_novos)} registros novos encontrados:")
        print(df_novos[chave])
    else:
        print("✅ Nenhum novo registro identificado.")

    # Comparar alterações em campos específicos
    df_comp = df_novo.merge(df_antigo, on=chave, how="inner", suffixes=("_novo", "_antigo"))
    mudancas = []

    for col in colunas_comparacao:
        col_novo = col + "_novo"
        col_antigo = col + "_antigo"

        if col_novo in df_comp.columns and col_antigo in df_comp.columns:
            df_dif = df_comp[df_comp[col_novo] != df_comp[col_antigo]]
            if not df_dif.empty:
                mudancas.append((col, df_dif[["Instrumento nº", col_novo, col_antigo]]))

    if mudancas:
        print("\n🔁 [ALTERAÇÕES DETECTADAS]")
        for col, df_part in mudancas:
            print(f" - Campo alterado: '{col}'")
            print(df_part.to_string(index=False))
    else:
        print("✅ Nenhuma alteração relevante nas colunas monitoradas.\n")
def main():
    print("🚀 Iniciando o robô Selenium...")

    navegador = conectar_navegador_existente()
    df_entrada = ler_planilha_entrada()

    if df_entrada.empty or "Instrumento nº" not in df_entrada.columns:
        print("❌ A planilha não contém dados válidos na coluna 'Instrumento nº'. Encerrando...")
        return

    try:
        saida_path = r"C:\Users\diego.brito\Downloads\robov1\Recursos\Resultado_Recursos.xlsx"
        base_antiga_path = r"C:\Users\diego.brito\Downloads\robov1\Recursos\Resultado_Recursos_Antigo.xlsx"

        if os.path.exists(saida_path):
            shutil.copyfile(saida_path, base_antiga_path)
            print(f"📂 Backup criado em: {base_antiga_path}")
        else:
            print("ℹ️ Nenhum arquivo anterior encontrado. Será a primeira coleta.")

        resultados_totais = []

        for idx, linha in df_entrada.iterrows():
            numero_instrumento = linha["Instrumento nº"]
            print(f"🔎 Buscando instrumento: {numero_instrumento}")

            sucesso = navegar_para_instrumento(navegador, numero_instrumento)

            if sucesso:
                print(f"✅ Instrumento {numero_instrumento} acessado com sucesso!")
                dados = verificar_e_registrar_repasses(navegador, numero_instrumento)
                if dados:
                    print(f"📊 Dados coletados para o instrumento {numero_instrumento}: {dados}")
                    resultados_totais.extend(dados)
                    registrar_excel(saida_path, dados)
                else:
                    print(f"⚠️ Nenhum dado coletado para o instrumento {numero_instrumento}.")
            else:
                print(f"⚠️ Não foi possível acessar o instrumento {numero_instrumento}.")
                registrar_excel(saida_path, [{
                    "Instrumento": numero_instrumento,
                    "Status": "Instrumento não encontrado"
                }])

            if idx < len(df_entrada) - 1:
                navegador.get("https://discricionarias.transferegov.sistema.gov.br/voluntarias/Principal/Principal.do")
                print(f"🔙 Retornando ao menu principal para o próximo instrumento.")
                # Aguardar um elemento conhecido do menu principal para confirmar carregamento
                esperar_elemento_por_xpath(navegador, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]")

        if resultados_totais:
            df_novo = pd.DataFrame(resultados_totais)
            df_novo = df_novo.rename(columns={"Instrumento": "Instrumento nº", "Valor Repassado": "Valor"})
            comparar_resultados(df_novo, base_antiga_path)
        else:
            print("⚠️ Nenhum dado coletado para comparação.")

    finally:
        navegador.quit()
        print("🔌 Navegador fechado.")

    print("🏁 Finalizado!")