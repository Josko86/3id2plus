from django.shortcuts import render
from django.http import HttpResponse
import templates
# Create your views here.
from scripts.obtenfci import obtenerFCI

def obtenerNumeroComanda(request):
    obtenerFCI()
    return render(request, 'obtenerfci.html')