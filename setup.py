# -*- coding: utf-8 -*-
from distutils.core import setup
import py2exe

setup(
    name="Buscar Imagenes",
    version="1.0",
    description="Busca imagenes que esten en un archivo excel y la copia a una carpeta",
    author="autor",
    scripts=["buscar y copiar imagenes version final.py"],
    console=["buscar y copiar imagenes version final.py"],
    options={"py2exe": {"bundle_files": 1}},
    zipfile=None,
)