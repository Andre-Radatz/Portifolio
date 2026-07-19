import requests
import pandas as pd
from bs4 import BeautifulSoup

# URL do site
url = 'https://books.toscrape.com/'

resposta = requests.get(url)
resposta.encoding = 'utf-8'  # Garante a decodificação correta

site = BeautifulSoup(resposta.text, 'html.parser')
info = site.find_all('article', class_='product_pod')

for produto in info:
    title = produto.find('h3').find('a').get('title')
    print(title)
    price = produto.find('p', class_='price_color').text
    print(price)
    stock = produto.find('p', class_='instock availability').text.strip()
    print(stock)
    print('---')

dados_livros = []
for produto in info:
    title = produto.find('h3').find('a').get('title')
    price = produto.find('p', class_='price_color').text
    stock = produto.find('p', class_='instock availability').text.strip()
    dados_livros.append({'Titulo': title, 'Preco': price, 'Estoque': stock})

# Cria o DataFrame
df_livros = pd.DataFrame(dados_livros)

print("Primeiros registros:")
print(df_livros.head())

# Exporta para CSV
df_livros.to_csv('livros.csv', index=False, encoding='utf-8')
print('Dados exportados para livros.csv')

print("\nDataFrame completo:")
print(df_livros)