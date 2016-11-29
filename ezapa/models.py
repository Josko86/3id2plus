from django.db import models
from django.utils import timezone


class Trabajador(models.Model):

    name = models.CharField(max_length=200)
    sueldo = models.CharField(max_length=200)

    def __str__(self):
        return self.name


class Dosier(models.Model):

    nombre = models.CharField(max_length=200)
    trabajador = models.ForeignKey(Trabajador, on_delete=models.CASCADE)
    tiempo = models.CharField(max_length=200)
    fecha_entrega = models.DateTimeField('date published', default=timezone.now())
    fecha_inicio = models.DateTimeField('date started', default=timezone.now())

    def __str__(self):
        return self.nombre
