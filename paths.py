import os
import sys


def base_dir():
    """Diretorio base da aplicacao: pasta do executavel quando compilado com
    PyInstaller, ou a pasta do projeto quando rodando via script."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))
