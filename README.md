# Automação de Upload de Documentos no SEI

## Sobre o Projeto
Este projeto tem como objetivo automatizar o envio de documentos no **SEI (Sistema Eletrônico de Informações)**, utilizando **Python** e **Selenium WebDriver**.  

O script faz a leitura de uma planilha Excel contendo processos, localiza os arquivos correspondentes em diretórios locais e realiza o upload automaticamente no SEI. Para evitar duplicações, cada envio é registrado em um arquivo de log no formato JSON.

---

## Funcionalidades
- Conexão ao navegador Google Chrome já aberto em modo debugger.
- Leitura de processos a partir de uma planilha Excel.
- Localização de subpastas com base no número do instrumento vinculado ao processo.
- Upload automático de arquivos PDF e ZIP.
- Preenchimento automático de campos obrigatórios:  
  - Data de elaboração  
  - Tipo de documento (Anexo)  
  - Visibilidade (Público)  
  - Marcação de nato-digital
- Registro em log JSON para evitar reenvio de arquivos.

---

## Estrutura do Projeto
.
├── script.py # Script principal de automação
├── upload_log.json # Registro dos arquivos já enviados
├── requirements.txt # Dependências do projeto
└── README.md # Documentação do projeto

yaml
Copiar
Editar

---

## Requisitos
- Python 3.8 ou superior
- Google Chrome instalado
- Dependências listadas em `requirements.txt`:
  - pandas  
  - selenium  
  - webdriver-manager  

Instale as dependências com:
```bash
pip install -r requirements.txt
Execução
Inicie o Google Chrome em modo debugger:

bash
Copiar
Editar
chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebugProfile"
Configure no script:

Caminho da planilha Excel (caminho_excel)

Caminho da pasta base dos documentos (caminho_base)

Execute o script:

bash
Copiar
Editar
python script.py
Log de Envio
O arquivo upload_log.json armazena todos os arquivos já enviados, garantindo que não sejam reenviados em execuções futuras.

Exemplo:

json
Copiar
Editar
{
  "12345.67890/2024-11": [
    "documento1.pdf",
    "documento2.zip"
  ]
}
Melhorias Futuras
Suporte a outros tipos de documentos além de "Anexo".

Geração de relatórios detalhados de sucesso e falha.

Substituição do log JSON por banco de dados relacional.

Interface gráfica para facilitar a configuração e execução.

Autor
Desenvolvido por Diego Brito.
