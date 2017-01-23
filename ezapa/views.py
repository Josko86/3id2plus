from django.shortcuts import render, render_to_response
from django.http import HttpResponse
from django.template import Context, Template
import templates
from scripts.ipon import ejecutar_ipon
from scripts.obtenfci import obtenerFCI

def obtenerNumeroComanda(request):

    result = obtenerFCI()
    c = Context({'result': result})
    return render(request, 'obtenerfci.html', c)

def ipon(request, nra):
    person = {
        '1': 'SUS se ha completado correctamente',
        '2': 'DDSD1  Ha tenido un error'
    }
    c = Context({'person': person})
    ejecutar_ipon(nra)
    return render(request, 'obtenerfci.html', c)

def prueba(request, nra):
    person = {
        '1' : 'SUS se ha completado correctamente',
        '2' : 'DDSD1  Ha tenido un error'
    }
    c = Context({'person': person})

    return render(request, 'obtenerfci.html', c)