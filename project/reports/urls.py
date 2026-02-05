from django.urls import path
from . import views

urlpatterns = [
    path('svod/', views.svod_report_page, name='svod_report_page'),
    path('svod/excel/', views.export_svod_excel, name='export_svod_excel'),
]