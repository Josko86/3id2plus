# -*- coding: utf-8 -*-
# Generated by Django 1.10.3 on 2016-11-25 11:35
from __future__ import unicode_literals

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Dosier',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('nombre', models.CharField(max_length=200)),
                ('tiempo', models.IntegerField(default=0)),
            ],
        ),
        migrations.CreateModel(
            name='Trabajador',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=200)),
                ('sueldo', models.IntegerField(default=0)),
            ],
        ),
        migrations.AddField(
            model_name='dosier',
            name='trabajador',
            field=models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='ezapa.Trabajador'),
        ),
    ]