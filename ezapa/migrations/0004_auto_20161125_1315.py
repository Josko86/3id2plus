# -*- coding: utf-8 -*-
# Generated by Django 1.10.3 on 2016-11-25 12:15
from __future__ import unicode_literals

import datetime
from django.db import migrations, models
from django.utils.timezone import utc


class Migration(migrations.Migration):

    dependencies = [
        ('ezapa', '0003_auto_20161125_1308'),
    ]

    operations = [
        migrations.AlterField(
            model_name='dosier',
            name='fecha_entrega',
            field=models.DateTimeField(default=datetime.datetime(2016, 11, 25, 12, 15, 46, 397789, tzinfo=utc), verbose_name='date published'),
        ),
        migrations.AlterField(
            model_name='dosier',
            name='fecha_inicio',
            field=models.DateTimeField(default=datetime.datetime(2016, 11, 25, 12, 15, 46, 397789, tzinfo=utc), verbose_name='date started'),
        ),
    ]
