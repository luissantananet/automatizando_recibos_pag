# Automatizando Recibos Pagamentos

Um aplicação desenvolvido em Python para coletar dados de um arquivo .xlsx e criar o .docx com os dados e texto pré-definidos.



## Instalação

1. Clone o repositório.
2. Instale as dependências com `pip install -r requirements.txt`.
3. Execute o aplicativo com `python Gerar_Recibos.py`.

## Dependências

- pandas 
- docx

## Como Usar

1. Abra o arquivo Excel com os dados dos colaboradores.
2. Preencha o arquivo Excel com os dados dos colaboradores.
3. Execute o aplicativo com `python Gerar_Recibos.py`.

## Criar .exe

1. Instale as dependências com `pip install pyinstaller`.
2. Execute o comando `pyinstaller --onefile Gerar_Recibos.py`.
3. O arquivo .exe será criado na pasta `dist`.

