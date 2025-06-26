# Automação de Upload de Documentos no SEI

Este projeto realiza a automação do envio de documentos externos (PDF e ZIP) para processos cadastrados no SEI (Sistema Eletrônico de Informações), utilizando o navegador Google Chrome em modo de depuração.

---

## Funcionalidades

- Conexão com instância do Chrome em execução via porta de depuração
- Leitura de planilha Excel com números de processo e instrumentos
- Busca automatizada no SEI e criação de documentos do tipo **Anexo**
- Upload automático de arquivos `.pdf` e `.zip` localizados em subpastas
- Registro de envios no arquivo `upload_log.json` para evitar duplicidade
- Logs detalhados com informações de progresso e erros

---

## Pré-requisitos

- Python 3.8 ou superior
- Google Chrome instalado
- Instalar bibliotecas Python:

```bash
pip install selenium pandas openpyxl webdriver-manager




Estrutura Esperada
Planilha Excel
A planilha deve conter as seguintes colunas:

Processo SEI (nº)

Instrumento nº

Caminho de exemplo:

python
Copiar
Editar
caminho_excel = r"C:\Users\seu.usuario\Downloads\planilha.xlsx"
Diretório Base
As subpastas devem conter no nome o número do instrumento, e dentro delas os arquivos .pdf ou .zip:

python
Copiar
Editar
caminho_base = r"C:\Users\seu.usuario\Documentos\ambiente_de_testes"
Como Executar
Inicie o Chrome com depuração remota:

bash
Copiar
Editar
chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebugProfile"
Execute o script principal:

bash
Copiar
Editar
python seu_script.py
Estrutura do Código
Conexão com navegador existente
python
Copiar
Editar
def conectar_navegador_existente(porta: int = 9222):
    opcoes = Options()
    opcoes.add_experimental_option("debuggerAddress", f"localhost:{porta}")
    return webdriver.Chrome(options=opcoes)
Log de envios
python
Copiar
Editar
def carregar_log_envio():
    if os.path.exists("upload_log.json"):
        with open("upload_log.json", "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def salvar_log_envio(log):
    with open("upload_log.json", "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)
Loop principal
python
Copiar
Editar
df = pd.read_excel(caminho_excel)
lista_processos = df['Processo SEI (nº)'].dropna().tolist()
log_envios = carregar_log_envio()

for processo in lista_processos:
    # Busca no SEI, cria documento, faz upload e atualiza log
    ...
Exemplo de upload_log.json
json
Copiar
Editar
{
  "1234567": [
    "documento1.pdf",
    "anexo_final.zip"
  ],
  "7654321": [
    "outro_arquivo.pdf"
  ]
}
Observações
O script depende da estabilidade dos seletores CSS/IDs da interface do SEI.

É altamente recomendável rodar o script inicialmente em ambiente de testes.

Arquivos já enviados são ignorados com base no log salvo em upload_log.json.

Licença
Este projeto é de uso interno e não possui uma licença pública definida.

Autor
Desenvolvido por Diego Bruno Santos de Brito.

yaml
Copiar
Editar

---
