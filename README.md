# PROJETO PYTHON - VENV/OPENPYXL - Carrega dados de um arquivo .xlsx (Excel), acessa células individuais ou por índice numérico, modifica e imprime na tela.

Este pequeno projeto tem o intuito de demonstrar a utilização do ambiente virtual
VENV e o módulo OPENPYXL.

VENV: Um ambiente virtual em Python isola dependências do projeto, evitando conflitos com pacotes globais do sistema. Ele permite que cada projeto tenha suas próprias bibliotecas e versões específicas.

OPENPYXL: módulo que permite a criação, leitura e edição de arquivos Excel (.xlsx e .xlsm) diretamente no Python. Ele é muito útil para automatizar relatórios, processar dados e criar planilhas dinâmicas.

Como dito no primeiro parágrafo, o projeto é bem simples e resume-se a explicar a utilização do ambiente virtual VENV e implementa um pequeno algoritmo que Carrega dados de um arquivo .xlsx (Excel), acessa células individuais ou por índice numérico, modifica e imprime na tela, com o módulo OPENPYXL.

## FERRAMENTAS UTILIZADAS
* Linguagem de programação Python.
* Ambiente virtual VENV.
* Git/GitHub
* Visual studio code.
* Windows 10.

## MODO DE UTILIZAR
* Clonar repositório.
* No diretório 'python-venv-openpyxl02', executar ```python -m venv venv``` para instalar o ambiente virtual.
* Executar ```pip install -r requirements.txt``` para instalar as dependências.
* Executar, caso esteja no Windows, ```venv\Scripts\activate``` para iniciar o ambiente. Caso Linux ou MacOS, ```source venv/bin/activate```.
* ```python app.py``` - Executa o algoritmo.
* Para sair do ambiente virtual ```deactivate```.

## COMANDOS IMPORTANTES
* ```python -m venv venv``` - Cria um ambiente virtual chamado venv. Observação: o primeiro venv é o comando, o segundo, o nome do diretório.
* No Windows, ```venv\Scripts\activate``` e no Linux, ```source venv/bin/activate``` - Inicializa o ambiente.
* ```deactivate``` - Encerra o ambiente.
* ```pip freeze > requirements.txt``` - Gera o arquivo que contém dependências. Esse mesmo comando atualiza o arquivo.
* ```pip list``` - Lista as dependências do projeto.
* ```pip show``` - Inserindo o nome da dependência após o comando, lista informações da dependência.
* ```pip install -r requirements.txt``` - Instala dependências que estão no arquivo 'requirements.txt'.
* ```pip install``` - Inserindo o nome da dependência após o comando, instala dependências.
* ```pip uninstall``` - Inserindo o nome da dependência após o comando, desinstala dependências.
