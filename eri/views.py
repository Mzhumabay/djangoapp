from django.shortcuts import render
from django.http import HttpResponse
import xml.etree.ElementTree as ET
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document
from django.db import connection
from django.conf import settings
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from io import BytesIO

def get_publications_data():
    with connection.cursor() as cursor:
        cursor.execute("SELECT * FROM publications")
        rows = cursor.fetchall()
    return rows

def get_info_data():
    with connection.cursor() as cursor:
        cursor.execute("SELECT * FROM info")
        rows = cursor.fetchall()
    return rows

def index(request):   
    return render(request, 'index.html')

def table_view(request):
    publications_data = get_publications_data()
    info_data = get_info_data()

    # Передаем данные в шаблон для отображения
    context = {
        'publications_data': publications_data,
        'info_data': info_data,
    }

    return render(request, 'table_view.html', context)


def download_docx(request):
    publications_data = get_publications_data()
    info_data = get_info_data()
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = 'attachment; filename="table_data.docx"'

    doc = Document()

    # Ваш код для создания DOCX документа с данными из таблиц
    doc.add_heading('Таблица публикаций:', level=2)
    for publication in publications_data:
        doc.add_paragraph(f"ID: {publication[0]}, Название: {publication[1]}")

    doc.add_heading('Таблица информации:', level=2)
    for info in info_data:
        doc.add_paragraph(f"ID: {info[0]}, Дата: {info[1]}, Описание: {info[2]}, ID Публикации: {info[3]}")

    doc.save(response)
    return response

def subsection_detail(request, subsection_id):
    # Получите содержание для подраздела на основе subsection_id
    # Для примера предположим, что у нас есть словарь с содержанием
    subsections = {
        'vvp': 'Валовый внутренний продукт (ВВП) - это...',
        'ifo': 'Индекс физического объема (ИФО) - это...',
        'prod_truda': 'Производительность труда - это...',
        'invest_osn_kapital': 'Инвестиции в основной капитал - это...',
        'zanyatost_bezrabotica': 'Занятость, безработица, средняя зарплата по основным странам - это...',
        'potreb_ceni_proizv_ceni': 'Индекс потребительских цен и индекс цен производителей - это...',
        'potreb_ceni_socialno_znac_tovary': 'Индекс цен на социально-значимые потребительские товары - это...',
        'potreb_ceni_po_stranam': 'Индекс потребительских цен по странам - это...',
        'mezhdunar_rezervy_kursy_valut': 'Международные резервы и курсы валют - это...',
        'gosdolg_v_pp_k_vvp_po_stranam': 'Госдолг в % к ВВП по странам - это...',
        'kreditny_reiting': 'Кредитный рейтинг - это...',
        'ispolnenie_gos_bjudzheta_dohody': 'Исполнение государственного бюджета (доходы) - это...',
        'ispolnenie_gos_bjudzheta_zatrati_deficit': 'Исполнение государственного бюджета (затраты и дефицит) - это...',
        'index_pmi_sovokupnyy_po_stranam': 'Индекс PMI (совокупный) по странам - это...',
        'index_pmi_promyshlennost_uslugi_po_stranam': 'Индекс PMI (в промышленности и услугах) по странам - это...',
        'torgovy_oborot_kazahstan': 'Торговый оборот Республики Казахстан - это...',
        'eksport_kazahstan_po_stranam': 'Экспорт Республики Казахстан в разрезе стран - это...',
        'import_kazahstan_po_stranam': 'Импорт Республики Казахстан в разрезе стран - это...',
        'osnovnye_eksportnye_tovary': 'Основные экспортные товары - это...',
        'osnovnye_importnye_tovary': 'Основные импортные товары - это...',
        'prognoz_socialno_ekonomich_razvitie': 'Прогноз социально-экономического развития - это...',
        'prognoz_institut_ekonomicheskih_issledovanij': 'Прогноз Института экономических исследований - это...',
        'konsensus_prognoz': 'Консенсус прогноз - это...',
        # Добавьте другие подразделы и их содержание здесь
    }
    
    content = subsections.get(subsection_id, None)
    
    if content is None:
        return render(request, 'not_found.html', {'subsection_id': subsection_id})

    return render(request, 'subsection_detail.html', {'content': content, 'subsection_id': subsection_id})
