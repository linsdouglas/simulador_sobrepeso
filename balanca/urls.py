from django.urls import path
from . import views

urlpatterns = [
    path('analise/', views.analise_ocorrencias, name='analise_ocorrencias'),
]
