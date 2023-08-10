from django.shortcuts import render
from django.http import HttpResponse
from docx import Document
from django.db import connection
import pandas as pd
import win32com.client as win32
from openpyxl import load_workbook
from io import BytesIO
import os


def get_publications_data():
    with connection.cursor() as cursor:
        cursor.execute("SELECT * FROM gdp_regions")
        rows = cursor.fetchall()
    return rows


def get_data(id):

    with connection.cursor() as cursor:
        cursor.execute(id)
        rows = cursor.fetchall()
        list_list = [list(tpl) for tpl in rows]
    return list_list


def index(request):
    return render(request, 'index.html')


def table_view(request):
    publications_data = get_publications_data()

    # Передаем данные в шаблон для отображения
    context = {
        'publications_data': publications_data
    }

    return render(request, 'table_view.html', context)


def excel_to_doc(request):
    # Загружаем Excel-файл с данными
    excel_file_path = 'static/doc/test_table.xlsx'
    wb = load_workbook(excel_file_path)
    ws = wb.active
# Создаем список списков для хранения данных
    data_list = []

# Проходим по строкам и столбцам таблицы и добавляем данные в список
    for row in ws.iter_rows():
        row_data = [cell.value for cell in row]
        data_list.append(row_data)

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = 'attachment; filename="table_data.docx"'

    doc = Document()

    # Ваш код для создания DOCX документа с данными из таблиц
    doc.add_heading('Таблица публикаций:', level=2)
    for publication in data_list:
        doc.add_paragraph(
            f"ID: {publication[0]}, Название: {publication[1]}, Название: {publication[2]}, Название: {publication[3]}")

    doc.save(response)
    return response


def download_excel(request):
    excel_file_path = 'static/doc/test_table.xlsx'  # Путь к вашему Excel файлу
    if os.path.exists(excel_file_path):
        with open(excel_file_path, 'rb') as excel_file:
            response = HttpResponse(excel_file.read(
            ), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response[
                'Content-Disposition'] = f'attachment; filename={os.path.basename(excel_file_path)}'
            return response
    else:
        # Обработка случая, если файл не найден
        return HttpResponse('Excel файл не найден', status=404)


def download_docx(request):
    publications_data = get_publications_data()
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = 'attachment; filename="table_data.docx"'

    doc = Document()

    # Ваш код для создания DOCX документа с данными из таблиц
    doc.add_heading('Таблица публикаций:', level=2)
    for publication in publications_data:
        doc.add_paragraph(f"ID: {publication[0]}, Название: {publication[1]}")

    doc.save(response)
    return response


def subsection_detail(request, subsection_id):
    excel_file_path = 'static/doc/test_table.xlsx'

    # Прочитать данные из Excel файла
    df = pd.read_excel(excel_file_path)

    # Преобразовать DataFrame в HTML-таблицу
    html_table = df.to_html(index=False, classes='table table-bordered')

    return render(request, 'excel_to_html.html', {'html_table': html_table})
#     # Получите содержание для подраздела на основе subsection_id
#     # Для примера предположим, что у нас есть словарь с содержанием

#     #content = subsections.get(subsection_id, None)


# # Получение записи и работы со словарем сырых запросов
#     raw_query_dict = RawQueryDictionary.objects.get(id=1)
#     raw_queries = raw_query_dict.raw_queries

# # Извлечение сырого запроса по ключу
#     query_id_1 = raw_queries.get("prod_truda")
#     #query_id_2 = raw_queries.get("query_id_2")

# # Обновление сырого запроса
#    # raw_queries["query_id_1"] = "UPDATE your_table SET column = value WHERE condition"
#     #raw_query_dict.save()
#     content  = get_data(query_id_1)
#     if content is None:
#         return render(request, 'not_found.html', {'subsection_id': subsection_id})

#     return render(request, 'subsection_detail.html', {'content': content, 'subsection_id': subsection_id})
