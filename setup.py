import os

os.environ['TCL_LIBRARY'] = "C:\\LOCAL_TO_PYTHON\\Python35-32\\tcl\\tcl8.6"
os.environ['TK_LIBRARY'] = "C:\\LOCAL_TO_PYTHON\\Python35-32\\tcl\\tk8.6"

from cx_Freeze import setup, Executable

setup(name = "audioSummarizer" ,
      version = "0.1" ,
      description = "" ,
      executables = [Executable("final.py")])