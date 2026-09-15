import os
import subprocess
import time

import psutil


def matar_chromedriver():
    """Mata todos os processos chromedriver.exe do usuario atual."""
    mortos = 0
    for proc in psutil.process_iter(['pid', 'name', 'username']):
        try:
            nome = (proc.info.get('name') or '').lower()
            if nome != 'chromedriver.exe':
                continue
            try:
                proc.kill()
                mortos += 1
            except Exception:
                pass
        except Exception:
            continue
    if mortos:
        time.sleep(1)
    return mortos


def matar_chrome_robot():
    """Mata processos chrome.exe do usuario atual."""
    mortos = 0
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            nome = (proc.info.get('name') or '').lower()
            if nome != 'chrome.exe':
                continue
            try:
                proc.kill()
                mortos += 1
            except Exception:
                pass
        except Exception:
            continue
    if mortos:
        time.sleep(1)
    return mortos


def matar_processo_chrome():
    """Limpa chromedriver e chrome gerados pelo robo."""
    try:
        matar_chromedriver()
    except Exception:
        pass
    try:
        matar_chrome_robot()
    except Exception:
        pass
    # Fallback via taskkill (caso psutil falhe)
    try:
        subprocess.run(['taskkill', '/IM', 'chromedriver.exe', '/F'],
                       capture_output=True, shell=True, timeout=10)
    except Exception:
        pass
    try:
        subprocess.run(['taskkill', '/IM', 'chrome.exe', '/F'],
                       capture_output=True, shell=True, timeout=10)
    except Exception:
        pass


def limpar_na_inicializacao():
    """Remove sobras de chromedriver/chrome do robo ao abrir o exe."""
    try:
        matar_chromedriver()
    except Exception:
        pass