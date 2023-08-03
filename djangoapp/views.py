from django.shortcuts import render

def index(request):
    # Здесь вы можете добавить код для извлечения данных из базы данных,
    # например, список таблиц из базы данных
    # table_list = Table.objects.all()  # Пример кода для извлечения списка таблиц (предполагается, что у вас есть модель "Table")

    # Для примера, создадим список таблиц с фиктивными данными
    table_list = ['Таблица 1', 'Таблица 2', 'Таблица 3', 'Таблица 4']

    context = {
        'table_list': table_list,
    }
    return render(request, 'index.html', context)