from django.urls import path
from . import views

urlpatterns = [
    path('analise/', views.analise_ocorrencias, name='analise_ocorrencias'),
    path('api/expedicao/', views.receber_expedicao, name='receber_expedicao'),
]

