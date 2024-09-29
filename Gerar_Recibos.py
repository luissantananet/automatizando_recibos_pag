import pandas as pd
from docx import Document
from docx.shared import Pt
import os

# Carrega a lista de colaboradores do arquivo Excel
colaboradores = pd.read_excel('colaboradores.xlsx')

# Nome do arquivo de saída
nome_arquivo = 'RECIBO DE PAGAMENTO DOS DOMINGOS GRAVATAI.docx'

# Verifica se o arquivo já existe e remove se necessário
if os.path.exists(nome_arquivo):
    os.remove(nome_arquivo)

# Cria um novo documento Word
document = Document()

# Define o tamanho da fonte para os campos
font_size = Pt(12)  # Tamanho padrão do texto
bold_font_size = Pt(14)  # Tamanho da fonte para campos em negrito

# Adiciona o texto do recibo para cada colaborador na mesma página
for _, row in colaboradores.iterrows():
    nome = row['nome']
    cpf = str(row['cpf'])  # Converte o CPF para string
    valor = str(row['valor'])  # Converte o valor para string
    
    p = document.add_paragraph()
    run = p.add_run("RECIBO DE PAGAMENTO\n")
    
    run = p.add_run(f"{nome}, ")
    run.bold = True
    run.font.size = bold_font_size
    
    run = p.add_run("inscrito(a) no CPF sob o nº ")
    
    run = p.add_run(cpf)
    run.bold = True
    run.font.size = bold_font_size
    
    run = p.add_run(", declaro para os devidos fins ter recebido nesta data, da empresa KFP SERVICE DIGITAL LTDA, inscrita no CNPJ sob o nº 41.230.154/0001-57, a importância de R$")
    
    run = p.add_run(valor)
    run.bold = True
    run.font.size = bold_font_size
    
    run = p.add_run(" concernente ao pagamento de um domingo trabalhado.\n\nCachoeirinha, 29 de setembro de 2024.\n\n_________________________________________________\nAssinatura\n")
    
    document.add_paragraph('\n')

# Salva o arquivo de recibos
document.save(nome_arquivo)

"""
CMD:
usar a ferramenta PyInstaller para transformar seu código Python em um executável (.exe). 
>>Agora, navegue até o diretório onde seu script Python está localizado e execute o seguinte comando:
pyinstaller --onefile seu_script.py
>>Substitua seu_script.py pelo nome do seu arquivo Python. A opção --onefile faz com que todos os arquivos sejam agrupados em um único executável.
>>Depois que o PyInstaller terminar, você encontrará seu arquivo .exe na pasta dist."""