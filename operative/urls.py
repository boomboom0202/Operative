from django.urls import path
from . import views

urlpatterns = [
    path('', views.StatisticsView.as_view(), name=''),
    path('statistics/excel/', views.export_statistics_excel, name='export_excel'),
]