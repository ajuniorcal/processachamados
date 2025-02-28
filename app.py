import streamlit as st
import requests
import json
import os
import sys
import pandas as pd
import re  

# ======================================
# ============ CONFIG LOGIN ============
# ======================================
USUARIO_PERMITIDO = "antonio.amaral"
SENHA_CORRETA = "Mudar@123"

# ======================================
# ============ CONFIG DESK =============
# ======================================
URL_AUTH = "https://api.desk.ms/Login/autenticar"
URL_BUSCAR_CHAMADO = "https://api.desk.ms/TablesMaestro1/lista"
URL_CRIAR_CHAMADO = "https://api.desk.ms/TablesMaestro1"

AUTH_HEADER = {"Authorization": "094186d770944a0427432958b8db4da55071cd8f"}
PUBLIC_KEY = {"PublicKey": "75e4ec2e5f76aa8b8f61a23f10c3f53b2106f169"}

# ======================================
# ============ CSS / STYLING ===========
# ======================================
st.markdown(
    """
    <style>
    /* Forçar o selectbox a abrir para baixo */
    .stSelectbox div[role="listbox"] {
        position: absolute !important;
        top: 100% !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ======================================
# ============ FUNÇÕES GERAIS ==========
# ======================================
def log_console(msg: str):
    """Força o log a aparecer no console do Streamlit."""
    print(msg)
    sys.stdout.flush()

def extrair_numero(nome: str) -> int:
    """Extrai números do nome do arquivo para ordenação correta."""
    match = re.search(r"(\d+)", nome)
    return int(match.group(1)) if match else float('inf')

def extrair_id_chamado_similar(valor: str) -> str:
    """Extrai o ID de 'Chamados similares' entre 'id=' e '&lang'."""
    match = re.search(r"id=([\w-]+)&lang", str(valor))
    return match.group(1) if match else ""

# ======================================
# ============ FUNÇÕES DESK ============
# ======================================
def autenticar():
    """Autentica na API Desk Manager e retorna um token."""
    log_console("[Auth] Solicitando novo token...")
    try:
        response = requests.post(URL_AUTH, headers=AUTH_HEADER, json=PUBLIC_KEY)
        response.raise_for_status()
        token = response.text.strip()
        token = json.loads(token)
        log_console(f"[Auth] Token gerado: {token}")
        return token
    except requests.RequestException as e:
        log_console(f"[Auth] Erro: {str(e)}")
        return None

def buscar_chamado_existente(token, numero_chamado: str) -> bool or None:
    """Verifica se um chamado já existe na Desk Manager."""
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

def criar_chamado(token, row: pd.Series) -> str or None:
    """Cria um novo chamado na Desk Manager."""
    headers = {"Authorization": token}
    # Corrigir valores NaT e None para strings vazias
    row = row.replace({pd.NaT: "", None: ""})
    row = row.astype(str)
    dados_chamado = {
        "TTableMaestro": {
            "Chave": "",
            "Campo26": row.get("Nível", ""),
            "Campo27": row.get("Andamento", ""),
            "Campo28": row.get("Número do chamado", ""),
            "Campo29": row.get("Nome do solicitante", ""),
            "Campo30": row.get("Filial (Solicitante)", ""),
            "Campo31": row.get("Tipo de chamado", ""),
            "Campo32": row.get("Status", ""),
            "Campo33": row.get("Operador", ""),
            "Campo34": row.get("Resolvido(a)s", ""),
            "Campo35": row.get("Fechado(a)s", ""),
            "Campo36": row.get("Data alvo", ""),
            "Campo37": row.get("Impacto", ""),
            "Campo38": row.get("Categoria", ""),
            "Campo39": row.get("Subcategoria", ""),
            "Campo40": row.get("Data alvo do SLA", ""),
            "Campo41": row.get("Data alvo de resposta", ""),
            "Campo42": row.get("Respondido", ""),
            "Campo43": row.get("Data de resposta", ""),
            "Campo44": row.get("Grupo de operadores", ""),
            "Campo45": row.get("Operador do escalonamento/rebaixamento", ""),
            "Campo46": row.get("Pedido", ""),
            "Campo47": row.get("Ação", ""),
            "Campo48": row.get("Anexos", ""),
            "Campo49": row.get("Chamados similares", ""),
        }
    }
    log_console(f"[Criação] Criando chamado {row.get('Número do chamado', 'SemNumero')}...")
    try:
        response = requests.put(URL_CRIAR_CHAMADO, headers=headers, json=dados_chamado)
        response.raise_for_status()
        log_console(f"[Criação] Sucesso -> ID: {response.text.strip()}")
        return response.text.strip()
    except requests.RequestException as e:
        log_console(f"[Criação] Erro: {str(e)}")
        return None

def processar_e_salvar_chamados(df: pd.DataFrame, nome_arquivo: str):
    """
    Processa a planilha e salva os chamados na Desk Manager.
    Retorna: total_inseridos (int), total_existentes (int), lista_de_erros (list)
    """
    log_console("[DeskManager] Iniciando processamento...")
    token = autenticar()
    if not token:
        log_console("[DeskManager] Erro ao autenticar.")
        return 0, 0, ["Erro na autenticação"]

    total_inseridos = 0
    total_existentes = 0
    erros = []

    # Processa todas as linhas da planilha (sem limitação)
    for index, row in df.iterrows():
        # Renovar token a cada 200 registros
        if index > 0 and index % 200 == 0:
            log_console("[Auth] Renovando token...")
            token = autenticar()
            if not token:
                log_console("[Auth] Falha na renovação do token.")
                break

        numero_chamado = row.get("Número do chamado", "")
        existe = buscar_chamado_existente(token, numero_chamado)
        if existe is None:
            erros.append(numero_chamado)
            continue
        if existe:
            log_console(f"[Info] Chamado {numero_chamado} já existe.")
            total_existentes += 1
            continue

        resultado = criar_chamado(token, row)
        if resultado:
            log_console(f"[Sucesso] Chamado {numero_chamado} criado.")
            total_inseridos += 1
        else:
            erros.append(numero_chamado)

    log_console("[Finalizado] Processamento completo.")
    return total_inseridos, total_existentes, erros

# ======================================
# ============ LOGIN SYSTEM ============
# ======================================
def login():
    """Exibe o formulário de login e verifica credenciais."""
    st.title("🔐 Login")
    usuario = st.text_input("Usuário", value="", max_chars=50)
    senha = st.text_input("Senha", type="password", value="", max_chars=50)
    if st.button("Entrar"):
        if usuario == USUARIO_PERMITIDO and senha == SENHA_CORRETA:
            st.session_state.logado = True
            st.rerun()  # Recarrega a página para ir à tela principal
        else:
            st.error("❌ Usuário ou senha incorretos!")

# ======================================
# ============ CHECK LOGIN =============
# ======================================
if "logado" not in st.session_state:
    st.session_state.logado = False

if not st.session_state.logado:
    login()
    st.stop()

# ======================================
# ============ LAYOUT PRINCIPAL =========
# ======================================
st.title("Processador de Planilhas - Desk Manager")

# Listar planilhas
planilhas_disponiveis = sorted(
    [f for f in os.listdir("planilhas") if f.endswith(".xlsx") and not f.startswith("~")],
    key=extrair_numero
)

if "planilha_df" not in st.session_state:
    st.session_state.planilha_df = None
if "planilha_selecionada" not in st.session_state:
    st.session_state.planilha_selecionada = None

# ======================================
# ============ CARREGAR PLANILHA =======
# ======================================
if st.session_state.planilha_df is None:
    planilha_selecionada = st.selectbox("Selecione uma planilha para processar:", planilhas_disponiveis)
    if st.button("Carregar Planilha"):
        log_console("[Streamlit] Botão 'Carregar Planilha' pressionado.")
        caminho_arquivo = os.path.join("planilhas", planilha_selecionada)
        df = pd.read_excel(caminho_arquivo)
        df.index = df.index + 1
        df.fillna("", inplace=True)
        if "Chamados similares" in df.columns:
            df["Chamados similares"] = df["Chamados similares"].apply(extrair_id_chamado_similar)
        st.session_state.planilha_df = df
        st.session_state.planilha_selecionada = planilha_selecionada
        st.success(f"Planilha '{planilha_selecionada}' carregada com sucesso!")
        st.dataframe(df.head())

# ======================================
# ============ EXECUTAR PLANILHA =======
# ======================================
if st.session_state.planilha_df is not None:
    if st.button("Executar Planilha"):
        st.info("Processando os dados... Aguarde.")
        log_console("[Streamlit] Botão 'Executar Planilha' pressionado.")
        total_inseridos, total_existentes, erros = processar_e_salvar_chamados(
            st.session_state.planilha_df,
            st.session_state.planilha_selecionada
        )
        st.success(f"{total_inseridos} Chamados inseridos com sucesso.")
        st.warning(f"{total_existentes} Chamados não inseridos porque já existiam.")
        if erros:
            st.error("Chamados não inseridos por erro:")
            st.write(erros)
