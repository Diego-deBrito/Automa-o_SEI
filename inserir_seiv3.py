# -*- coding: utf-8 -*-
"""
Este script automatiza o processo de upload de documentos para o sistema SEI
(Sistema Eletrônico de Informações).

Funcionalidades Principais:
1.  Lê uma planilha Excel para obter uma lista de números de processo SEI e
    seus respectivos números de instrumento.
2.  Conecta-se a uma sessão do Google Chrome já aberta em modo de depuração para
    interagir com uma sessão autenticada do SEI.
3.  Para cada processo, localiza uma pasta correspondente no sistema de arquivos local
    baseado no número do instrumento.
4.  Itera sobre todos os arquivos PDF e ZIP dentro da pasta encontrada.
5.  Para cada arquivo, o robô navega no SEI, abre o formulário de "Documento Externo",
    preenche os metadados (tipo "Anexo", data atual, nato-digital, público) e
    realiza o upload do arquivo.
6.  Mantém um log em formato JSON (`upload_log.json`) para rastrear os arquivos
    já enviados e evitar uploads duplicados.

Pré-requisitos:
- Python 3.x
- Bibliotecas: pandas, selenium, webdriver-manager
- Google Chrome instalado.
- Uma instância do Chrome deve ser iniciada com a depuração remota ativada.
  Exemplo de comando (Windows):
  "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebug"
"""

import os
import time
import glob
import json
import pandas as pd
from datetime import date
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException

# --- Bloco de Configuração ---
# Centraliza caminhos e parâmetros para fácil manutenção.
CONFIG = {
    "caminho_excel": r"C:\Users\diego.brito\Downloads\teste_sei\RTMA Passivo 2024 - PROJETOS E PROGRAMAS 1.xlsx",
    "caminho_base_documentos": r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Documentos\ambiente de testes",
    "arquivo_log": "upload_log.json",
    "porta_debugger": 9222,
    "timeout_geral": 20  # Tempo máximo de espera para elementos da página em segundos
}

# Centraliza os seletores do SEI para facilitar atualizações caso a interface mude.
SEI_SELECTORS = {
    "campo_pesquisa_rapida": "txtPesquisaRapida",
    "iframe_visualizacao": "ifrVisualizacao",
    "link_novo_documento": "#divArvoreAcoes a[href*='controlador.php?acao=documento_escolher_tipo']",
    "tabela_tipos_documento": "#tblSeries > tbody > tr:nth-child(1)",
    "dropdown_tipo_anexo": "#selSerie",
    "campo_data_elaboracao": "txtDataElaboracao",
    "checkbox_nato_digital": "lblNato",
    "checkbox_publico": "lblPublico",
    "campo_nome_arquivo": "txtNumero",
    "campo_upload_arquivo": "filArquivo",
    "botao_salvar": "btnSalvar",
    "iframe_progresso_upload": "ifrProgressofrmAnexos"
}

# --- Funções Auxiliares ---

def conectar_navegador_existente(porta):
    """
    Conecta a uma instância do Google Chrome em execução na porta de depuração especificada.

    Args:
        porta (int): O número da porta onde o modo de depuração do Chrome está rodando.

    Returns:
        webdriver.Chrome or None: Retorna o objeto do driver do navegador se a conexão
                                  for bem-sucedida, caso contrário, retorna None.
    """
    try:
        print(f"[INFO] Tentando conectar ao navegador na porta {porta}...")
        opcoes_chrome = Options()
        opcoes_chrome.add_experimental_option("debuggerAddress", f"localhost:{porta}")
        navegador = webdriver.Chrome(options=opcoes_chrome)
        print("[SUCCESS] Conectado ao navegador existente com sucesso!")
        return navegador
    except WebDriverException:
        print("[ERROR] Falha ao conectar. Verifique se o Chrome está aberto com o modo de depuração ativado.")
        print(f'          Exemplo: chrome.exe --remote-debugging-port={porta} --user-data-dir="C:\\ChromeDebugProfile"')
        return None
    except Exception as e:
        print(f"[ERROR] Ocorreu um erro inesperado ao tentar conectar ao navegador: {e}")
        return None

def carregar_log_envio(caminho_arquivo):
    """
    Carrega o log de arquivos já enviados a partir de um arquivo JSON.

    Args:
        caminho_arquivo (str): O caminho para o arquivo de log.

    Returns:
        dict: Um dicionário com o histórico de uploads. Retorna um dicionário
              vazio se o arquivo não existir.
    """
    if os.path.exists(caminho_arquivo):
        with open(caminho_arquivo, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def salvar_log_envio(log_data, caminho_arquivo):
    """
    Salva o dicionário de log de envios em um arquivo JSON.

    Args:
        log_data (dict): O dicionário contendo o log atualizado.
        caminho_arquivo (str): O caminho onde o arquivo de log será salvo.
    """
    with open(caminho_arquivo, "w", encoding="utf-8") as f:
        json.dump(log_data, f, ensure_ascii=False, indent=4)

def navegar_e_preparar_formulario(navegador, wait, processo_sei):
    """
    Realiza a sequência de navegação no SEI para chegar até o formulário de upload
    de documento externo e preenche os campos comuns.

    Args:
        navegador (webdriver.Chrome): A instância do driver do Selenium.
        wait (WebDriverWait): A instância do objeto de espera explícita.
        processo_sei (str): O número do processo a ser buscado.

    Returns:
        bool: True se a preparação foi bem-sucedida, False caso contrário.
    """
    try:
        # Garante que o foco está no conteúdo principal antes de cada nova ação.
        navegador.switch_to.default_content()

        # Realiza a busca pelo processo
        campo_busca = wait.until(EC.presence_of_element_located((By.ID, SEI_SELECTORS["campo_pesquisa_rapida"])))
        campo_busca.clear()
        campo_busca.send_keys(str(processo_sei))
        campo_busca.send_keys(Keys.ENTER)
        print("[INFO] Pesquisa pelo processo enviada.")

        # O conteúdo principal do processo está dentro de um iframe. É necessário mudar o contexto.
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, SEI_SELECTORS["iframe_visualizacao"])))
        print("[INFO] Contexto alterado para o iframe de visualização.")

        # Clica no ícone para incluir um novo documento
        botao_novo_doc = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SEI_SELECTORS["link_novo_documento"])))
        botao_novo_doc.click()
        print("[INFO] Navegando para a tela de escolha de tipo de documento.")

        # Na nova tela, seleciona o tipo "Documento Externo"
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SEI_SELECTORS["tabela_tipos_documento"]))).click()
        print("[INFO] Tipo 'Documento Externo' selecionado.")

        # Preenche os campos padrão do formulário
        dropdown_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SEI_SELECTORS["dropdown_tipo_anexo"])))
        Select(dropdown_element).select_by_visible_text("Anexo")
        print("[INFO] Tipo de documento definido como 'Anexo'.")

        data_atual = date.today().strftime("%d/%m/%Y")
        campo_data = navegador.find_element(By.ID, SEI_SELECTORS["campo_data_elaboracao"])
        campo_data.clear()
        campo_data.send_keys(data_atual)
        print(f"[INFO] Data de elaboração preenchida com: {data_atual}")

        navegador.find_element(By.ID, SEI_SELECTORS["checkbox_nato_digital"]).click()
        print("[INFO] Documento marcado como 'Nato-Digital'.")

        navegador.find_element(By.ID, SEI_SELECTORS["checkbox_publico"]).click()
        print("[INFO] Nível de acesso definido como 'Público'.")

        return True

    except TimeoutException:
        print("[ERROR] Tempo de espera excedido ao tentar localizar um elemento durante a navegação.")
        return False
    except Exception as e:
        print(f"[ERROR] Erro inesperado durante a preparação do formulário: {e}")
        return False

# --- Script Principal ---

def main():
    """
    Orquestra todo o processo de automação.
    """
    navegador = conectar_navegador_existente(CONFIG["porta_debugger"])
    if not navegador:
        print("[ERROR] Não foi possível conectar ao navegador. Encerrando o script.")
        return

    try:
        df = pd.read_excel(CONFIG["caminho_excel"])
        # Garante que a coluna de processos não tenha valores nulos antes de converter para lista
        lista_processos = df['Processo SEI (nº)'].dropna().astype(str).tolist()
        print(f"[INFO] Planilha carregada. {len(lista_processos)} processos encontrados.")
    except FileNotFoundError:
        print(f"[ERROR] O arquivo Excel não foi encontrado no caminho: {CONFIG['caminho_excel']}")
        return
    except Exception as e:
        print(f"[ERROR] Falha ao ler o arquivo Excel: {e}")
        return

    log_envios = carregar_log_envio(CONFIG["arquivo_log"])
    wait = WebDriverWait(navegador, CONFIG["timeout_geral"])

    # Itera sobre cada processo da planilha
    for processo in lista_processos:
        try:
            print("-" * 60)
            print(f"[INFO] Iniciando automação para o processo: {processo}")

            # Inicializa o log para o processo se ele for novo
            if processo not in log_envios:
                log_envios[processo] = []

            # Encontra o número do instrumento associado ao processo no DataFrame
            instrumento_match = df.loc[df['Processo SEI (nº)'] == processo, 'Instrumento nº']
            if instrumento_match.empty:
                print(f"[WARNING] Número do instrumento não encontrado para o processo {processo}. Pulando.")
                continue

            numero_instrumento = str(instrumento_match.values[0])
            print(f"[INFO] Procurando pastas de documentos para o instrumento: {numero_instrumento}")

            # Busca pastas no diretório base que contenham o número do instrumento
            subpastas = [p for p in os.listdir(CONFIG["caminho_base_documentos"]) if os.path.isdir(os.path.join(CONFIG["caminho_base_documentos"], p))]
            pastas_compativeis = [p for p in subpastas if numero_instrumento in p]

            if not pastas_compativeis:
                print(f"[WARNING] Nenhuma pasta de documentos encontrada contendo '{numero_instrumento}'.")
                continue

            # Itera sobre cada pasta de documento encontrada
            for pasta in pastas_compativeis:
                caminho_pasta = os.path.join(CONFIG["caminho_base_documentos"], pasta)
                print(f"[INFO] Analisando a pasta: {caminho_pasta}")

                # Busca recursivamente por arquivos PDF e ZIP dentro da pasta
                arquivos_para_upload = sorted([
                    f for f in glob.glob(os.path.join(caminho_pasta, '**', '*.*'), recursive=True)
                    if f.lower().endswith(('.pdf', '.zip'))
                ])

                if not arquivos_para_upload:
                    print(f"[INFO] Nenhum arquivo .pdf ou .zip encontrado em: {caminho_pasta}")
                    continue

                # Itera sobre cada arquivo a ser enviado
                for caminho_completo_arquivo in arquivos_para_upload:
                    nome_arquivo = os.path.basename(caminho_completo_arquivo)

                    if nome_arquivo in log_envios[processo]:
                        print(f"[WARNING] O arquivo '{nome_arquivo}' já foi enviado anteriormente para este processo. Pulando.")
                        continue

                    print(f"\n[INFO] Preparando para enviar o arquivo: {nome_arquivo}")

                    # A cada novo arquivo, refaz a navegação para garantir que a interface está no estado esperado.
                    if not navegar_e_preparar_formulario(navegador, wait, processo):
                        print("[ERROR] Falha ao preparar o formulário. Pulando para o próximo arquivo.")
                        continue

                    # Preenche os campos específicos do arquivo e realiza o upload
                    try:
                        campo_nome = wait.until(EC.presence_of_element_located((By.ID, SEI_SELECTORS["campo_nome_arquivo"])))
                        campo_nome.clear()
                        campo_nome.send_keys(nome_arquivo)
                        print(f"[INFO] Nome do anexo preenchido com: '{nome_arquivo}'")

                        campo_upload = navegador.find_element(By.ID, SEI_SELECTORS["campo_upload_arquivo"])
                        campo_upload.send_keys(caminho_completo_arquivo)
                        print(f"[INFO] Arquivo '{caminho_completo_arquivo}' selecionado para upload.")

                        navegador.find_element(By.ID, SEI_SELECTORS["botao_salvar"]).click()
                        print("[INFO] Botão 'Confirmar Dados' clicado. Aguardando finalização do upload...")

                        # Aguarda o desaparecimento da barra de progresso, indicando que o upload terminou.
                        wait.until(EC.invisibility_of_element_located((By.ID, SEI_SELECTORS["iframe_progresso_upload"])))
                        print(f"[SUCCESS] Upload do arquivo '{nome_arquivo}' finalizado com sucesso.")

                        # Registra o sucesso no log
                        log_envios[processo].append(nome_arquivo)
                        salvar_log_envio(log_envios, CONFIG["arquivo_log"])
                        print("[INFO] Log de envio atualizado.")

                    except Exception as erro_upload:
                        print(f"[ERROR] Falha durante o processo de upload do arquivo '{nome_arquivo}'.")
                        print(f"          Detalhes: {erro_upload}")
                        # Garante que o robô volte ao estado inicial para a próxima tentativa
                        navegador.switch_to.default_content()
                        continue

        except Exception as erro_processo:
            print(f"[ERROR] Ocorreu um erro geral e inesperado ao processar o processo {processo}.")
            print(f"          Detalhes: {erro_processo}")
            # Tenta se recuperar voltando ao contexto padrão do navegador
            navegador.switch_to.default_content()
            continue

    print("-" * 60)
    print("[SUCCESS] Todos os processos da planilha foram verificados.")


if __name__ == "__main__":
    main()
