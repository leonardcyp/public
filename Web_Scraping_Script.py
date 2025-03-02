from bs4 import BeautifulSoup
import requests
import pandas as pd


for i in range(1,50):

    url = f'https://books.toscrape.com/catalogue/category/books_1/page-{i}.html'

    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    order_list = soup.find('ol', class_='row')
    articles= soup.findAll('article', class_='product_pod')

    books=[]
    for article in articles:
        img= article.find('img')
        title= img.attrs['alt']
        star= article.find('p')
        star=star['class'][1]
        price=article.find('p', class_='price_color').text
        price=float(price[2:])
        books.append([title, star, price])

df = pd.DataFrame(books, columns=['Title', 'Star', 'Price'])
df.to_csv('books.csv')