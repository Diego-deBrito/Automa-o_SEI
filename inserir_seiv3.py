import os
import time
import glob
import pandas as pd

from datetime import date
from selenium import webdriver
from selenium.common import StaleElementReferenceException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager






# --- Função para conectar ao navegador já aberto com modo debugger ---
def conectar_navegador_existente(porta: int = 9222):
    """
    Conecta a uma instância do Google Chrome já em execução na porta de depuração especificada.
    """
    try:
        print(f"Tentando conectar ao navegador na porta {porta}...")
        opcoes_chrome = Options()
        opcoes_chrome.add_experimental_option("debuggerAddress", f"localhost:{porta}")
        navegador = webdriver.Chrome(options=opcoes_chrome)
        print(" X.X Conectado ao navegador existente com sucesso!")
        return navegador
    except WebDriverException:
        print(" Erro ao conectar. Verifique se o Chrome está aberto com depuração:")
        print(f'chrome.exe --remote-debugging-port={porta} --user-data-dir="C:\\ChromeDebugProfile"')
        return None
    except Exception as e:
        print(f"Erro inesperado: {e}")
        return None


import json

LOG_ARQUIVOS = "upload_log.json"

def carregar_log_envio():
    if os.path.exists(LOG_ARQUIVOS):
        with open(LOG_ARQUIVOS, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def salvar_log_envio(log):
    with open(LOG_ARQUIVOS, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)

# --- INÍCIO SCRIPT PRINCIPAL ---


# Caminho do Excel atualizado
caminho_excel = r"C:\Users\diego.brito\Downloads\teste_sei\RTMA Passivo 2024 - PROJETOS E PROGRAMAS 1.xlsx"

# Caminho base onde estão os documentos
caminho_base = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Documentos\ambiente de testes"

# Conecta ao navegador já aberto
navegador_conectado = conectar_navegador_existente()


if not navegador_conectado:
    print("⚠ Navegador não conectado. Encerrando script.")
    exit()

# Carrega os dados do Excel
df = pd.read_excel(caminho_excel)
lista_processos = df['Processo SEI (nº)'].dropna().tolist()  # remove vazios

print(f" Total de processos encontrados: {len(lista_processos)}")


log_envios = carregar_log_envio()

# Loop principal de processos
for processo in lista_processos:
    try:
        print(f"\n🟡 Iniciando automação para o processo: {processo}")
        wait = WebDriverWait(navegador_conectado, 15)

        if processo not in log_envios:
            log_envios[processo] = []


        # Localiza e interage com o campo de busca
        campo_busca = wait.until(EC.presence_of_element_located((By.ID, "txtPesquisaRapida")))
        campo_busca.clear()
        campo_busca.send_keys(str(processo))
        campo_busca.send_keys(Keys.ENTER)
        print(" Pesquisa enviada.")

        # Acessa o iframe onde está a árvore de ações
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao")))
        print(" Foco alterado para o iframe.")

        # Clica no botão de novo documento externo
        seletor_link = "#divArvoreAcoes a[href*='controlador.php?acao=documento_escolher_tipo']"
        botao_novo_doc = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, seletor_link)))
        botao_novo_doc.click()
        print(" Clique no botão 'Novo Documento' feito com sucesso.")

        # Espera o botão externo e clica nele
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "infraTrClara")))
        btn_externo_css = "#tblSeries > tbody > tr:nth-child(1)"
        navegador_conectado.find_element(By.CSS_SELECTOR, btn_externo_css).click()
        print(" Documento externo selecionado.")

        # Seleciona o tipo de documento como "Anexo"
        dropdown_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#selSerie")))
        select = Select(dropdown_element)
        select.select_by_visible_text("Anexo")
        print(" Tipo de documento 'Anexo' selecionado.")

        time.sleep(3)
        # Preenche data com data atual
        data_atual = date.today().strftime("%d/%m/%Y")
        campo_data = navegador_conectado.find_element(By.ID, "txtDataElaboracao")
        campo_data.clear()
        campo_data.send_keys(data_atual)
        print(f" Data preenchida: {data_atual}")

        # Marca como nato-digital
        navegador_conectado.find_element(By.ID, "lblNato").click()
        print(" Marcado como Nato-Digital.")

        # Marca como público
        navegador_conectado.find_element(By.ID, "lblPublico").click()
        print(" Marcado como Público.")

        # zzzVerifica o número do instrumento no DataFrame
        instrumento_match = df.loc[df['Processo SEI (nº)'] == processo, 'Instrumento nº']
        if instrumento_match.empty:
            print(" Instrumento não encontrado no Excel.")
            navegador_conectado.switch_to.default_content()
            continue

        numero_instrumento = str(instrumento_match.values[0])
        print(f" Procurando subpastas com número do instrumento: {numero_instrumento}")

        # Lista subpastas que contenham o número do instrumento
        subpastas = [p for p in os.listdir(caminho_base) if os.path.isdir(os.path.join(caminho_base, p))]
        pastas_compativeis = [p for p in subpastas if numero_instrumento in p]

        if not pastas_compativeis:
            print(f"⚠ Nenhuma pasta com '{numero_instrumento}' encontrada.")
            navegador_conectado.switch_to.default_content()
            continue

        for pasta in pastas_compativeis:
            caminho_pasta = os.path.join(caminho_base, pasta)
            print(f" Explorando pasta: {caminho_pasta}")

            arquivos = sorted([
                f for f in glob.glob(os.path.join(caminho_pasta, '**', '*.*'), recursive=True)
                if f.lower().endswith(('.pdf', '.zip'))
            ])

            if not arquivos:
                print(f"⚠ Nenhum arquivo PDF ou ZIP em: {caminho_pasta}")
                continue

            for arquivo in arquivos:
                nome_arquivo = os.path.basename(arquivo)

                # Verifica se já foi enviado
                if nome_arquivo in log_envios[processo]:
                    print(f"⚠️ Arquivo já enviado anteriormente: {nome_arquivo} — pulando.")
                    continue

                try:
                    print(f"\n📄 Iniciando envio do arquivo: {nome_arquivo}")

                    # Garante que está no contexto principal antes de tudo
                    navegador_conectado.switch_to.default_content()

                    # 🔁 Refaz a busca do processo
                    campo_busca = wait.until(EC.presence_of_element_located((By.ID, "txtPesquisaRapida")))
                    campo_busca.clear()
                    campo_busca.send_keys(str(processo))
                    campo_busca.send_keys(Keys.ENTER)
                    print("🔍 Pesquisa enviada.")

                    # Acessa novamente o iframe com a árvore
                    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrVisualizacao")))
                    print("📥 Foco alterado para o iframe.")

                    # Clica no botão "Novo Documento"
                    seletor_link = "#divArvoreAcoes a[href*='controlador.php?acao=documento_escolher_tipo']"
                    botao_novo_doc = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, seletor_link)))
                    botao_novo_doc.click()
                    print("📄 Clique no botão 'Novo Documento' executado.")

                    # Espera o tipo "Documento Externo" e clica na primeira linha
                    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "infraTrClara")))
                    btn_externo_css = "#tblSeries > tbody > tr:nth-child(1)"
                    navegador_conectado.find_element(By.CSS_SELECTOR, btn_externo_css).click()
                    print("📄 Documento externo selecionado.")

                    # Seleciona o tipo "Anexo"
                    dropdown_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#selSerie")))
                    select = Select(dropdown_element)
                    select.select_by_visible_text("Anexo")
                    print("📄 Tipo de documento 'Anexo' selecionado.")

                    # Preenche a data atual no campo correspond
                    time.sleep(1)
                    data_atual = date.today().strftime("%d/%m/%Y")
                    campo_data = navegador_conectado.find_element(By.ID, "txtDataElaboracao")
                    campo_data.clear()
                    campo_data.send_keys(data_atual)
                    print(f"📅 Data preenchida com: {data_atual}")

                    # Marca como Nato-Digital
                    checkbox_nato = navegador_conectado.find_element(By.ID, "lblNato")
                    checkbox_nato.click()
                    print("✅ Marcado como Nato-Digital.")

                    # Marca como Público
                    checkbox_publico = navegador_conectado.find_element(By.ID, "lblPublico")
                    checkbox_publico.click()
                    print("✅ Marcado como Público.")

                    # Preenche o campo "txtNumero" com o nome do arquivo
                    try:
                        campo_numero = wait.until(EC.presence_of_element_located((By.ID, "txtNumero")))
                        campo_numero.clear()
                        campo_numero.send_keys(nome_arquivo)
                        print(f"📝 Campo 'txtNumero' preenchido com: {nome_arquivo}")
                    except Exception as e:
                        print(f"❌ Erro ao preencher o campo 'txtNumero': {e}")
                        continue  # Pula para o próximo arquivo se falhar

                    # Envia o arquivo para o campo de upload
                    campo_upload = wait.until(EC.presence_of_element_located((By.ID, "filArquivo")))
                    campo_upload.send_keys(arquivo)
                    print(f"📤 Arquivo enviado para o campo de upload: {arquivo}")

                    # Confirma os dados preenchidos
                    botao_confirmar = wait.until(EC.element_to_be_clickable((By.ID, "btnSalvar")))
                    botao_confirmar.click()
                    print("🆗 Clique em 'Confirmar Dados' executado.")

                    # Aguarda o iframe de progresso sumir, indicando fim do upload
                    WebDriverWait(navegador_conectado, 30).until(
                        EC.invisibility_of_element_located((By.ID, "ifrProgressofrmAnexos"))
                    )
                    print("✅ Upload finalizado com sucesso.")

                    # Salva o nome do arquivo no log de envios
                    log_envios[processo].append(nome_arquivo)
                    salvar_log_envio(log_envios)
                    print(f"📝 Arquivo registrado no log: {nome_arquivo}")

                except Exception as erro_arquivo:
                    print(f"❌ Erro ao processar o arquivo: {nome_arquivo}")
                    print(f"   ➤ Detalhes: {erro_arquivo}")
                    navegador_conectado.switch_to.default_content()
                    continue


    except Exception as e:
        print(f"❌ Erro geral no processo {processo}: {e}")
        navegador_conectado.switch_to.default_content()
        continue

print(" Todos os processos foram finalizados com sucesso! :)")
navegador_conectado.quit()

