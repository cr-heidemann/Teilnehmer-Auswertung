from distutils.core import setup
import py2exe
import os
import glob
import pandas as pd
from docxtpl import DocxTemplate
import tkinter  as tk
from Auswertung_gui import main 
import tkinter.filedialog as fd
import sys


setup(
    windows=[
        {
            "script": "gui.pyw",
            "icon_resources": [(1, "index.ico")]
        }],options={
        'py2exe': 
        {
            'includes': ['lxml.etree', 'lxml._elementpath', 'gzip'],
        }
    })

