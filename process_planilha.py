import os
import re
import pandas as pd
from html import unescape  # Para decodificar caracteres HTML como &ecirc; -> ê

# Obtendo o caminho absoluto da pasta do script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_PLANILHAS = os.path.join(BASE_DIR, "planilhas")

# Definição das colunas esperadas
COLUNAS_ESPERADAS = [
    'Nível', 'Andamento', 'Número do chamado', 'Nome do solicitante', 'Filial (Solicitante)', 
    'Tipo de chamado', 'Status', 'Operador', 'Resolvido(a)s', 'Fechado(a)s', 'Data alvo', 
    'Impacto', 'Categoria', 'Subcategoria', 'Data alvo do SLA', 'Data alvo de resposta', 
    'Respondido', 'Data de resposta', 'Grupo de operadores', 'Operador do escalonamento/rebaixamento', 
    'Pedido', 'Ação', 'Anexos', 'Chamados similares'
]

def extrair_numero(nome_arquivo):
    """Extrai o número do arquivo para ordenação correta."""
    numeros = re.findall(r'\d+', nome_arquivo)  # Captura números no nome do arquivo
    return int(numeros[0]) if numeros else 0  # Retorna o número como inteiro

def listar_planilhas():
    """Lista todos os arquivos .xlsx na pasta 'planilhas', ordenados corretamente"""
    if not os.path.exists(PASTA_PLANILHAS):
        return []  # Retorna lista vazia se a pasta não existir
    
    arquivos = [f for f in os.listdir(PASTA_PLANILHAS) if f.endswith(".xlsx")]
    
    # Ordenar numericamente, em vez de alfabeticamente
    arquivos.sort(key=extrair_numero)
    
    return arquivos

def extrair_id_chamado_similar(valor):
    """Extrai o ID correto da coluna 'Chamados similares'."""
    if isinstance(valor, str):
        match = re.search(r'id=([\w\d-]+)&?', valor)  # Captura o ID entre 'id=' e '&'
        return match.group(1) if match else None  # Retorna apenas o ID
    return None

def limpar_html(texto):
    """Remove tags HTML e decodifica entidades HTML."""
    if isinstance(texto, str):
        texto_limpo = re.sub(r'<[^>]+>', '', texto)  # Remove tags HTML
        return unescape(texto_limpo).strip()  # Decodifica entidades HTML (&ecirc; -> ê) e remove espaços extras
    return texto

def carregar_planilha(nome_arquivo):
    """Carrega a planilha, validando e tratando os dados."""
    caminho_arquivo = os.path.join(PASTA_PLANILHAS, nome_arquivo)
    
    try:
        df = pd.read_excel(caminho_arquivo, engine="openpyxl", dtype=str)  # Tudo será string

        # Verificar se todas as colunas esperadas estão presentes
        colunas_faltando = [col for col in COLUNAS_ESPERADAS if col not in df.columns]
        if colunas_faltando:
            return None, f"Erro: Faltam as colunas {colunas_faltando} na planilha {nome_arquivo}."
        
        # Ajustar a coluna 'Número do chamado' para substituir '\' por '-'
        if 'Número do chamado' in df.columns:
            df['Número do chamado'] = df['Número do chamado'].str.replace(re.escape("\\"), "-", regex=True)

        # Tratar a coluna 'Chamados similares' para extrair o ID correto
        if 'Chamados similares' in df.columns:
            df['Chamados similares'] = df['Chamados similares'].apply(extrair_id_chamado_similar)

        # Limpar HTML da coluna 'Anexos'
        if 'Anexos' in df.columns:
            df['Anexos'] = df['Anexos'].apply(limpar_html)

        return df, None  # Retorna o dataframe processado e nenhum erro
    
    except Exception as e:
        return None, f"Erro ao carregar a planilha {nome_arquivo}: {str(e)}"
