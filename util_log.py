import os
from paths import pasta_logs


def escrever_log(arquivo, linha):
    try:
        pasta = pasta_logs()
        os.makedirs(pasta, exist_ok=True)
        caminho = os.path.join(pasta, arquivo)
        with open(caminho, 'a', encoding='utf-8') as f:
            f.write(linha + '\n')
    except Exception:
        pass
