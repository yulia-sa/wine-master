import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas


WINERY_YEAR_OF_FOUNDATION = 1920
WINE_EXCEL = 'wine.xlsx'


def count_winery_age():
    return datetime.datetime.now().year - WINERY_YEAR_OF_FOUNDATION


def define_year_ending(year):
    year_str = str(year)
    if len(year_str) >= 2 and year_str[-2:] in ['11', '12', '13', '14']:
        return 'лет'
    if year_str[-1:] in ['0', '5', '6', '7', '8', '9']:
        return 'лет'
    if year_str[-1:] in ['2', '3', '4']:
        return 'года'
    if year_str[-1:] == '1':
        return 'год'
    return ''


def read_wines_excel(file_name):
    excel_data_df = pandas.read_excel(
        file_name,
        sheet_name='Лист1',
        usecols=[0, 1, 2, 3, 4, 5],
        na_values=['N/A', 'NA'],
        keep_default_na=False
    )

    wines_list_of_dicts = excel_data_df.to_dict(orient='records')
    wines_list_of_dicts_by_category = defaultdict(list)

    for wine_dict in wines_list_of_dicts:
        category = wine_dict.get('Категория')
        wines_list_of_dicts_by_category[category].append(wine_dict)

    return wines_list_of_dicts_by_category


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    winery_age = count_winery_age()
    years_format = define_year_ending(winery_age)
    wines_data_from_excel = read_wines_excel(WINE_EXCEL)

    rendered_page = template.render(
        winery_age=winery_age,
        years_format=years_format,
        wines_by_category=wines_data_from_excel,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
