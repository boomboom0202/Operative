from django.urls import path
from . import views

urlpatterns = [
    path('', views.StatisticsView.as_view(), name=''),
]