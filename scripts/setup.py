#!/usr/bin/python
# -*- coding: utf-8 -*-

from cx_Freeze import setup, Executable

setup(
    name="ipon",
    version="0.1",
    description="Ezapa",
    executables=[Executable("ipon.py")],
)