import os
from datetime import datetime
from paths import base_dir

PASTA_LOGS = os.path.join(base_dir(), 'logs')


def escrever_log(arquivo, linha):
    try:
        os.makedirs(PASTA_LOGS, exist_ok=True)
        caminho = os.path.join(PASTA_LOGS, arquivo)
        with open(caminho, 'a', encoding='utf-8') as f:
            f.write(linha + '\n')
    except Exception:
        pass
