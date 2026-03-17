import requests
import os
import shutil
from datetime import datetime
import zipfile
from pathlib import Path
from auto_sofia import executar_sofia_completo
import time as time_module
from calendar import monthrange

# Configuração
API_KEY = "minha_senha_muito_secreta"
VPS_URL = os.getenv('VPS_URL', 'http://10.88.172.116:8000')
HEADERS = {"X-API-KEY": "minha_senha_muito_secreta"}


def consultar_empresas(cliente_id):
    try:
        resp = requests.get(
            f'{VPS_URL}/api/sofia',
            headers=HEADERS,
            data={'cliente_id': cliente_id, 'acao':'consultar_empresas'},
            timeout=120
        )
        resp.raise_for_status()
        return resp.json()
    
    except requests.RequestException as e:
        print(f"[API] Erro ao consultar empresas: {e}")
        return None
    
def consultar_execucao(data_execucao):
    try:
        resp = requests.get(
            f'{VPS_URL}/api/sofia',
            headers=HEADERS,
            data={'exec_id': data_execucao, 'acao': "consultar_execucao"},
            timeout=120
        )
        resp.raise_for_status()
        return resp.json()
    except requests.RequestException as e:
        print(f"[API] Erro ao consultar execução: {e}")
        return None

def criar_execucao(empresa_id, per_inicial, per_final):
    try:
        resp = requests.post(
            f'{VPS_URL}/api/sofia',
            headers=HEADERS,
            data={
                'acao': 'criar_execucao',
                'empresa_id': empresa_id,
                'per_inicial': per_inicial,
                'per_final': per_final,
            },
            timeout=120
        )
        resp.raise_for_status()
        return resp.json().get('exec_id')
    except requests.RequestException as e:
        print(f"[API] Erro ao criar execução: {e}")
        return None

def atualizar_status_execucao(exec_id, status):
    try:
        resp = requests.post(
            f'{VPS_URL}/api/sofia',
            headers=HEADERS,
            data={'acao': 'atualizar_status_execucao', 'exec_id': exec_id, 'status': status},
            timeout=120
        )
        resp.raise_for_status()
        return True
    except requests.RequestException as e:
        print(f"[API] Erro ao atualizar status: {e}")
        return False
    

def upload_arquivos(zip_path, exec_id):
    try:    
        # Envia para VPS
        with open(zip_path, 'rb') as f:
            resp = requests.post(
                f'{VPS_URL}/api/sofia',
                headers=HEADERS,
                data={'exec_id': exec_id, 'acao': "upload_execucao"},
                files={'arquivo_zip': f},
                timeout=120
            )
        resp.raise_for_status()
        return True
    except requests.RequestException as e:
        print(f"[API] Erro ao fazer upload do arquivo: {e}")

def consultar_configuracao(): #agendado
    try:
        resp = requests.get(
            f'{VPS_URL}/api/sofia',
            headers=HEADERS,
            data={"hoje":datetime.now().day, "acao":"consultar_configuracao"},
            timeout=120
        )
        resp.raise_for_status()
        return resp.json()
    
    except requests.RequestException as e:
        print(f"[API] Erro ao consultar configuração: {e}")
        return None

def consultar_configuracao_cliente(cliente_id): #manual
    try:
        resp = requests.get(
            f"{VPS_URL}/api/sofia",
            headers=HEADERS,
            data={"cliente_id":cliente_id, "acao":"consultar_configuracao_cliente"},
            timeout=120
        )
        resp.raise_for_status()
        return resp.json()
    except requests.RequestException as e:
        print(f"[API] Erro ao consultar configuração: {e}")
        return None

def consultar_execucoes_pendentes():
    try:
        resp = requests.get(
            f'{VPS_URL}/api/sofia',
            headers=HEADERS,
            data={"acao": "consultar_execucoes_pendentes"},
            timeout=120
        )
        resp.raise_for_status()
        return resp.json()
    except requests.RequestException as e:
        print(f"[API] Erro ao consultar pendentes: {e}")
        return None

def verificar_execucao_existente(empresa_id, per_inicial, per_final):
    try:
        resp = requests.get(
            f'{VPS_URL}/api/sofia',
            headers=HEADERS,
            data={
                'acao': 'verificar_execucao_existente',
                'empresa_id': empresa_id,
                'per_inicial': per_inicial,
                'per_final': per_final,
            },
            timeout=120
        )
        resp.raise_for_status()
        return resp.json().get('existe', False)
    except requests.RequestException as e:
        print(f"[API] Erro ao verificar execução existente: {e}")
        return False

#senha: fiscal012026
#user: cla00005
if __name__ == "__main__":
    BASE_DOWNLOADS = r"E:\Projetos\4 - SOFIA\DOWNLOADS"
    TEMP_DIR = r"E:\Projetos\4 - SOFIA\TEMP"
    os.makedirs(BASE_DOWNLOADS, exist_ok=True)
    os.makedirs(TEMP_DIR, exist_ok=True)

    def calcular_periodo_mes_anterior():
        hoje = datetime.now()
        if hoje.month == 1:
            ano, mes = hoje.year - 1, 12
        else:
            ano, mes = hoje.year, hoje.month - 1
        _, ultimo_dia_mes = monthrange(ano, mes)
        return datetime(ano, mes, 1).strftime("%d/%m/%Y"), datetime(ano, mes, ultimo_dia_mes).strftime("%d/%m/%Y")

    # ← 1. Adicionado usuario_sefaz como parâmetro
    def processar_empresa(empresa, per_inicial, per_final, senha_sefaz, usuario_sefaz, exec_id=None):
        cnpj = empresa['cnpj']
        data_hoje = datetime.now().strftime("%Y%m%d")

        temp_empresa = os.path.join(TEMP_DIR, cnpj)
        os.makedirs(temp_empresa, exist_ok=True)

        if exec_id:
            atualizar_status_execucao(exec_id, "EXECUTANDO")

        dict_infos = {
            "cnpj": cnpj,
            "nfe1": empresa['nfe_emitente'],
            "nfe2": empresa['nfe_destinatario'],
            "nfce": empresa['nfce'],
            "cte1": empresa['cte_emitente'],
            "cte2": empresa['cte_tomador'],
            "faturamento": empresa['faturamento'],
        }

        print(f"Iniciando automação para {empresa['nome']}...")
        executar_sofia_completo(
            dict_infos=dict_infos,
            per_inicial=per_inicial,
            per_final=per_final,
            senha=senha_sefaz,
            usuario=usuario_sefaz,  # ← 2. Passando usuario
            download_dir=temp_empresa,
            headless=False
        )

        arquivos = [f for f in Path(temp_empresa).glob('*') if f.is_file()]

        if not arquivos:
            print(f"Nenhum arquivo baixado para {empresa['nome']}")
            shutil.rmtree(temp_empresa)
            if exec_id:
                atualizar_status_execucao(exec_id, "ERRO")
            return False

        if exec_id is None:
            exec_id = criar_execucao(empresa['id'], per_inicial, per_final)
            if not exec_id:
                print("Falha ao criar execução, pulando empresa")
                shutil.rmtree(temp_empresa)
                return False
            atualizar_status_execucao(exec_id, "EXECUTANDO")

        zip_path = os.path.join(BASE_DOWNLOADS, f"{cnpj}_{data_hoje}.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for arquivo in arquivos:
                zipf.write(arquivo, arquivo.name)

        shutil.rmtree(temp_empresa)
        print(f"ZIP criado: {zip_path}")

        sucesso = upload_arquivos(zip_path, exec_id)
        print(f"Upload {'concluído' if sucesso else 'FALHOU'} — exec_id={exec_id}")

        if not sucesso:
            atualizar_status_execucao(exec_id, "ERRO")

        return sucesso

    INTERVALO_SEGUNDOS = 60


    while True:
        try:
            # ---- CAMINHO 1: Requisições manuais pendentes ----
            pendentes = consultar_execucoes_pendentes()

            if pendentes:
                cliente_id = pendentes[0]['empresa']['cliente']
                configuracao_pendentes = consultar_configuracao_cliente(cliente_id)
                senha_sefaz = configuracao_pendentes["senha_sefaz"] if configuracao_pendentes else None
                usuario_sefaz = configuracao_pendentes["usuario_sefaz"] if configuracao_pendentes else None

                if not senha_sefaz or not usuario_sefaz:
                    print("Credenciais SEFAZ não encontradas para processar pendentes.")
                else:
                    for execucao in pendentes:
                        empresa = execucao['empresa']
                        per_ini = datetime.strptime(execucao['periodo_inicial'], "%Y-%m-%d").strftime("%d/%m/%Y")
                        per_fin = datetime.strptime(execucao['periodo_final'], "%Y-%m-%d").strftime("%d/%m/%Y")
                        print(f"\nPendente: {empresa['nome']} | {per_ini} → {per_fin}")
                        processar_empresa(empresa, per_ini, per_fin, senha_sefaz, usuario_sefaz, exec_id=execucao['id'])
            else:
                print("Nenhuma requisição manual pendente.")

            # ---- CAMINHO 2: Execução agendada ----
            print("Verificando execução agendada...")
            configuracao = consultar_configuracao()

            if not configuracao:
                print("Não é dia de execução agendada.")
            else:
                per_inicial, per_final = calcular_periodo_mes_anterior()  # ← 4. Dinâmico
                print(f"Período calculado: {per_inicial} → {per_final}")

                for config in configuracao:
                    cliente_id = config['cliente']
                    senha_sefaz = config['senha_sefaz']
                    usuario_sefaz = config['usuario_sefaz']
                    print(f"\nCliente {cliente_id}")

                    empresas = consultar_empresas(cliente_id=cliente_id)
                    if not empresas:
                        print(f"Nenhuma empresa para cliente {cliente_id}")
                        continue

                    per_ini_db = datetime.strptime(per_inicial, "%d/%m/%Y").strftime("%Y-%m-%d")
                    per_fin_db = datetime.strptime(per_final, "%d/%m/%Y").strftime("%Y-%m-%d")

                    for empresa in empresas:
                        print(f"\nEmpresa: {empresa['nome']} | CNPJ: {empresa['cnpj']}")

                        if verificar_execucao_existente(empresa['id'], per_ini_db, per_fin_db):
                            print(f"Já existe execução para esse período, pulando.")
                            continue

                        processar_empresa(empresa, per_inicial, per_final, senha_sefaz, usuario_sefaz)

        except Exception as e:
            print(f"[ERRO NO CICLO] {e}")

        print(f"\n[{datetime.now()}] Aguardando 5 minutos para próximo ciclo...")
        time_module.sleep(INTERVALO_SEGUNDOS)