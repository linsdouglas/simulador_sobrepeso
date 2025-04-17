from django.urls import path
from . import views
from .views import upload_sap

urlpatterns = [
    path('analise/', views.analise_ocorrencias, name='analise_ocorrencias'),
    path('api/expedicao/', views.receber_expedicao, name='receber_expedicao'),
    path('api/upload_sap/', upload_sap, name='upload_sap'),
]

