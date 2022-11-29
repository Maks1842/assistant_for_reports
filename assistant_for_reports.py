'''
С помощью данного модуля из файла EXCEL извлекаются данные
и через заданные теги вставляются в текст шаблона .DOCX
'''
import fake_useragent
import requests
from bs4 import BeautifulSoup

import re
import pandas
from docxtpl import DocxTemplate


HEADERS = {'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:107.0) Gecko/20100101 Firefox/107.0', 'accept': '*/*'}
HOST = ''
user = fake_useragent.UserAgent().random

header = {
    'user-agent': user
}

'''
1. Panda извлекает все данные из excel
2. По индексу/названию столбца формирует список из данных этого столбца
3. Индексам (которые обозначены в шаблоне .docx) присваиваются значения из списка
4. Данные по индексам передаются в шаблон .docx
5. Присваивается имя новому файлу и он сохраняется в указанной директории
'''
def insert_in_doc_pattern():

    excel_data_df = pandas.read_excel('data/Тест в json.xlsx')

    count_list = []
    data_dict = {}

    for x in range(330):
        count = x+1
        count_list.append(count)
        dataframe = excel_data_df[count].tolist()

        name = dataframe[2]
        date = dataframe[4]
        web = dataframe[5]
        num_person = dataframe[87]
        balls = dataframe[7]

        address = parse(web)

        # data_dict.update({
        #     f"{count}": {
        #         'name': name,
        #         'date': date,
        #         'web': web,
        #         'num_person': num_person,
        #         'balls': balls}
        # })

        context = {
                'name': f"{name}",
                'date': f"{date}",
                'address': f"{address}",
                'num_person': f"{num_person}",
                'balls': f"{balls}"
        }

        name_file = re.sub(r'\"', '', name)

        doc = DocxTemplate(f'data/Index шаблон отчета.docx')
        doc.render(context)
        doc.save(f"result/Отчеты ЧР_обр 2022/Отчет_{name_file}.docx")


'''
Извлекает информацию о юридическом адресе организации с официального сайта.
Важное условие: 
- сайт известен заранее и содержится в файле excel (с которым работает Panda);
- сайты всех проверяемых организаций идентичны, в противном случае возникает исключение.
'''
def parse(web):

    print(f'{web}')

    address = []
    try:
        html = get_html(f'https://{web}/frontpage/page/30')
        if html.status_code == 200:
            soup = BeautifulSoup(html.text, 'html.parser')

            items = soup.find_all('div', class_='data-front-list data-front-name')
            address = items[8].find('span', class_='descr-list').text.strip()
    except Exception:
        address = 'Необходимо вставить адрес'

    return address


def get_html(url, params=None):
    d = requests.get(url, headers=HEADERS, params=params)
    return d


insert_in_doc_pattern()