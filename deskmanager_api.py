import requests
import json
import os
import sys
import pandas as pd  # Para tratar valores NaN

# Configuração da API Desk Manager
URL_AUTH = "https://api.desk.ms/Login/autenticar"
URL_BUSCAR_CHAMADO = "https://api.desk.ms/TablesMaestro1/lista"
URL_CRIAR_CHAMADO = "https://api.desk.ms/TablesMaestro1"

AUTH_HEADER = {"Authorization": "094186d770944a0427432958b8db4da55071cd8f"}  
PUBLIC_KEY = {"PublicKey": "75e4ec2e5f76aa8b8f61a23f10c3f53b2106f169"}

def log_console(msg):
    """Força logs a aparecerem no console"""
    print(msg)
    sys.stdout.flush()

def autenticar():
    """Realiza a autenticação e retorna o token"""
    log_console("[Auth] Solicitando novo token...")
    try:
        response = requests.post(URL_AUTH, headers=AUTH_HEADER, json=PUBLIC_KEY)
        response.raise_for_status()
        token = response.text.strip()
        token = json.loads(token)  # Remove escapes
        log_console(f"[Auth] Token gerado: {token}")
        return token
    except requests.RequestException as e:
        log_console(f"[Auth] Erro: {str(e)}")
        return None

def buscar_chamado_existente(token, numero_chamado):
    """Verifica se o chamado já existe"""
    headers = {"Authorization": token}
    payload = {"Pesquisa": numero_chamado, "Ativo": "1"}

    log_console(f"[Busca] Verificando chamado {numero_chamado}...")
    try:
        response = requests.post(URL_BUSCAR_CHAMADO, headers=headers, json=payload)
        response.raise_for_status()
        data = response.json()
        log_console(f"[Busca] Retorno: {data}")

        return data.get("total") != "0"
    except requests.RequestException as e:
        log_console(f"[Busca] Erro: {str(e)}")
        return None

def criar_chamado(token, row):
    """Cria um novo chamado"""
    headers = {"Authorization": token}
    
    # Garantir que todos os valores vazios sejam substituídos por ""
    row = row.fillna("") if isinstance(row, pd.Series) else row

    dados_chamado = {
        "TTableMaestro": {
            "Chave": "",
            "Campo26": row["Nível"],
            "Campo27": row["Andamento"],
            "Campo28": row["Número do chamado"],
            "Campo29": row["Nome do solicitante"],
            "Campo30": row["Filial (Solicitante)"],
            "Campo31": row["Tipo de chamado"],
            "Campo32": row["Status"],
            "Campo33": row["Operador"],
            "Campo34": row["Resolvido(a)s"],
            "Campo35": row["Fechado(a)s"],
            "Campo36": row["Data alvo"],
            "Campo37": row["Impacto"],
            "Campo38": row["Categoria"],
            "Campo39": row["Subcategoria"],
            "Campo40": row["Data alvo do SLA"],
            "Campo41": row["Data alvo de resposta"],
            "Campo42": row["Respondido"],
            "Campo43": row["Data de resposta"],
            "Campo44": row["Grupo de operadores"],
            "Campo45": row["Operador do escalonamento/rebaixamento"],
            "Campo46": row["Pedido"],
            "Campo47": row["Ação"],
            "Campo48": row["Anexos"],
            "Campo49": row["Chamados similares"],
        }
    }

    log_console(f"[Criação] Criando chamado {row['Número do chamado']}...")
    try:
        response = requests.put(URL_CRIAR_CHAMADO, headers=headers, json=dados_chamado)
        response.raise_for_status()
        log_console(f"[Criação] Sucesso -> ID: {response.text.strip()}")
        return response.text.strip()
    except requests.RequestException as e:
        log_console(f"[Criação] Erro: {str(e)}")
        return None

def processar_e_salvar_chamados(df, nome_arquivo):
    """Processa e envia os chamados"""
    log_console("[DeskManager] Iniciando processamento...")
    token = autenticar()
    if not token:
        log_console("[DeskManager] Erro ao autenticar.")
        return 0, 0, ["Erro na autenticação"]

    total_inseridos = 0
    total_existentes = 0
    erros = []

    df = df.head(250)

    for index, row in df.iterrows():
        if index > 0 and index % 200 == 0:
            log_console("[Auth] Renovando token...")
            token = autenticar()
            if not token:
                log_console("[Auth] Falha na renovação.")
                break

        numero_chamado = row["Número do chamado"]
        if buscar_chamado_existente(token, numero_chamado):
            log_console(f"[Info] Chamado {numero_chamado} já existe.")
            total_existentes += 1
            continue

        if criar_chamado(token, row):
            log_console(f"[Sucesso] Chamado {numero_chamado} criado.")
            total_inseridos += 1
        else:
            erros.append(numero_chamado)

    log_console("[Finalizado] Processamento completo.")
    return total_inseridos, total_existentes, erros
