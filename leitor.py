import fitz  # PyMuPDF
import pandas as pd
import re

# Caminho para o arquivo PDF
caminho_pdf = 'C:/Users/gabriel.gomes/Desktop/leitor_pdf/12.pdf'

def extrair_dados(caminho_pdf):
    doc = fitz.open(caminho_pdf)
    texto = ''
    for pagina in doc:
        texto += pagina.get_text()
    return texto

# Extrair texto do PDF
texto = extrair_dados(caminho_pdf)


# Adicionar o regex para pegar o nome do procedimento
padrao = re.compile(r'Procedimento(.*?) - \n(.*?)\nFonte', re.DOTALL)
procedimentos = re.findall(padrao, texto)

# Dividir o texto em linhas
linhas = texto.split('\n')

quantidades = []
for i, linha in enumerate(linhas):
    if 'Quantidade' in linha:
        try:
            # Tentar extrair o número na linha após "Quantidade"
            quantidade = int(linhas[i + 2].strip())
            quantidades.append(quantidade)
        except ValueError:
            # Se não for possível converter o texto para int, ignorar
            pass
        except IndexError:
            # Se estivermos no final do arquivo e não houver mais linhas, ignorar
            pass


# Imprimir os procedimentos encontrados
for procedimento in procedimentos:
    print("Procedimento:", procedimento[1])  # Ajuste o índice conforme necessário

# Salvando os dados em uma tabela Excel
df = pd.DataFrame({'Procedimento': [procedimento[1] for procedimento in procedimentos],
                   'Quantidade': quantidades[:len(procedimentos)]})  # Ajuste conforme necessário

# Salvar o DataFrame como um arquivo Excel
df.to_excel('procedimentos_quantidades.xlsx', index=False)
