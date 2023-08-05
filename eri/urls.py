from django.urls import path, include
from eri import views

urlpatterns = [
    path('', views.index, name='index'),
    path('<str:subsection_id>/', views.subsection_detail, name='subsection_detail'),
    path('table/', views.table_view, name='table_view'),
    #path('download_pdf/', views.download_pdf, name='download_pdf'),
    path('download_docx/', views.download_docx, name='download_docx'),
    #path('download_xml/', views.download_xml, name='download_xml'),
]