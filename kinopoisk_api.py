import json
import requests
import pandas as pd
import urllib.parse

films = pd.read_excel(
    'file.xlsx',
    usecols = "B,E")
films.columns = ['title','year']

text_begin = 'https://api.kinopoisk.dev&name='
text_token = '&token=XXXXXXX-XXXXXXX-XXXXXXX-XXXXXXX'
text_year = '&year='
linktokino = 'https://www.kinopoisk.ru/film/'

df = pd.DataFrame(columns=['url', 'title', 'rating','year'])
for film_number in range(0,446):
    film = films.loc[film_number]
    film_year = str(film.year)
    encoded_string = urllib.parse.quote(film.title)
    text_full = text_begin+encoded_string+text_token+text_year+film_year
    response = requests.get(text_full)
    items = json.loads(response.text)
    if items['pages'] ==0:
        df_err = pd.DataFrame([[0,film.title,0,0]],columns=['url', 'title', 'rating','year'])
        df = pd.concat([df, df_err], ignore_index=True)
    else:
        item_id = str(items['docs'][0]['id'])
        item_name = items['docs'][0]['name']
        item_rating = items['docs'][0]['rating']['kp']
        item_year = items['docs'][0]['year']
        export = linktokino+item_id, item_name, item_rating, item_year
        df_true = pd.DataFrame([export], columns=['url', 'title', 'rating','year'])
        df = pd.concat([df, df_true], ignore_index=True)
df.to_excel("output.xlsx")