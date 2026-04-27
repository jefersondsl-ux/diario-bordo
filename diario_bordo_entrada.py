"""
Diário de Bordo — Entrada de Dados (Standalone)
================================================
App Streamlit único para cadastro/atualização de projetos e diário de bordo.
Funciona de forma independente em cada máquina do analista.
Os caminhos das bases são resolvidos automaticamente pelo usuário Windows.

CORREÇÕES APLICADAS (v2):
  - Bug 1: Arquivo temporário agora é criado FORA do OneDrive (tempfile.gettempdir())
            e copiado com shutil.copy2 — elimina o conflito de sync que gerava cópia
            com nome da máquina.
  - Bug 2: Lock compartilhado agora aguarda LOCK_SYNC_WAIT segundos após criação
            para garantir que o OneDrive sincronizou o arquivo antes de prosseguir.
            Também verifica propriedade do lock (maquina + usuario) após o wait.
  - Bug 3: salvar_controle recebe o mesmo tratamento de escrita segura via temp local.

MELHORIAS (v3) — Robustez do lock OneDrive:
  - Melhoria 1: Jitter aleatório (0–3 s) antes de tentar o lock para reduzir
                colisão simultânea entre analistas.
  - Melhoria 2: LOCK_SYNC_WAIT aumentado de 4 s → 10 s para cobrir conexões lentas.
  - Melhoria 3: Verificação dupla da propriedade do lock após o sync wait.
  - Melhoria 4: _limpar_conflitos_onedrive() remove arquivos de conflito gerados
                pelo OneDrive após cada gravação bem-sucedida.
"""

import os
import html as html_mod
import json
import random
import time
import shutil
import tempfile
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime, date
from pathlib import Path

# ==============================
# CAMINHOS DINÂMICOS (POR USUÁRIO)
# ==============================

_ONEDRIVE = os.path.join(os.path.expanduser("~"), "OneDrive - Claro SA")
_BASE_GOV  = os.path.join(_ONEDRIVE, "BASES", "Projetos_GOV")
_BD_DIM    = os.path.join(_BASE_GOV, "Diario_Bordo", "BD_DIM")
_BD_BORDO  = os.path.join(_BASE_GOV, "Diario_Bordo", "BD_Diario_Bordo")
_BD_SGP    = os.path.join(_BASE_GOV, "Base_Dados_SGP", "Bases_Processadas_Python")

PATH_CONTROLE   = os.path.join(_BD_DIM,   "d_Controle_Projetos.xlsx")
PATH_DIARIO     = os.path.join(_BD_BORDO, "f_Diario_Bordo.xlsx")
PATH_APONT      = os.path.join(_BD_DIM,   "d_apontamentos.xlsx")
PATH_RESP       = os.path.join(_BD_DIM,   "d_responsaveis.xlsx")
PATH_TECNOLOGIA = os.path.join(_BD_DIM,   "d_tecnologia.xlsx")
PATH_BACKLOG    = os.path.join(_BD_SGP,   "BD_Backlog_SGP.xlsx")

# ==============================
# LOCK — CONFIGURAÇÕES
# ==============================

# Lock compartilhado no OneDrive (visível a todos os PCs)
LOCK_DIARIO    = Path(PATH_DIARIO).with_suffix(".lock")
LOCK_CONTROLE  = Path(PATH_CONTROLE).with_suffix(".lock")

LOCK_TIMEOUT   = 30   # segundos aguardando lock ser liberado
LOCK_MAX_IDADE = 120  # segundos para considerar lock morto (crash/desligamento)
LOCK_SYNC_WAIT = 10   # segundos aguardando OneDrive sincronizar o lock antes de confirmar


# ==============================
# MAPA CEC → COORDENADOR
# ==============================

CEC_COORD_MAP = {
    "AMANDA RODRIGUES DE ALMEIDA":                   "KEYLLA LORRANNY",
    "ANA CAROLINA ALMEIDA RODRIGUES MAGALHAES":      "KEYLLA LORRANNY",
    "FLAVIA SILVA DANTAS":                           "KEYLLA LORRANNY",
    "QUEDMA DUARTE DA SILVA":                        "KEYLLA LORRANNY",
    "PHILIPPE DINIZ MELO":                           "KEYLLA LORRANNY",
    "EDUARDO DA SILVA SOUSA":                        "CARLOS ANDRÉ",
    "GEANDRE SIMPLICIO DO NASCIMENTO":               "CARLOS ANDRÉ",
    "DULCE TEIXEIRA NOVAES":                         "CARLOS ANDRÉ",
    "CAMILLA FRANCO FERREIRA":                       "CARLOS ANDRÉ",
    "BRUNO RIBEIRO BORGES":                          "CARLOS ANDRÉ",
    "HELLEN CRISTINA DO NASCIMENTO SILVA":           "CARLOS ANDRÉ",
    "RAICK NARDES DA SILVA":                         "CARLOS ANDRÉ",
    "RAUSELIZ DE SOUSA VIEIRA NASCIMENTO":           "CARLOS ANDRÉ",
    "SARAH CARDOSO BARROS":                          "CARLOS ANDRÉ",
    "DANIEL CARNEIRO DE CARVALHO":                   "CARLOS ANDRÉ",
    "ADAMOR MARTINS DE SOUSA":                       "MAURICÉIA ABE",
    "CRISTIANE DA SILVA AMORIM NEVES":               "MAURICÉIA ABE",
    "FRANCISCO JOSE LIMA DA SILVA":                  "MAURICÉIA ABE",
    "SONIA GERMINIA DA CONCEICAO NOGUEIRA":          "MAURICÉIA ABE",
    "CLAUDIO EVANGELISTA DA SILVA":                  "MAURICÉIA ABE",
    "ELAINE ANDRE DE SOUSA FERREIRA":                "MAURICÉIA ABE",
    "LAURO SERGIO MELO LEITE":                       "MAURICÉIA ABE",
    "CLEBER FERREIRA DE BARROS":                     "MAURO MAGALHÃES",
    "HIEDA CAPISTRANO DA PENHA DIAS":                "MAURO MAGALHÃES",
    "MAYALU FERREIRA NERY DOS SANTOS":               "MAURO MAGALHÃES",
    "VANESSA MARTINS ELOY":                          "MAURO MAGALHÃES",
    "RENATA COSTA DE BRITO":                         "MAURO MAGALHÃES",
    "MARCELO DE SOUZA RODRIGUES":                    "MAURO MAGALHÃES",
    "GLORIA ADRIANA DUARTE DOS SANTOS":              "MAURO MAGALHÃES",
}

# ==============================
# LOCK DE ESCRITA — CORRIGIDO (Bug 2)
# ==============================

def _maquina() -> str:
    return os.environ.get("COMPUTERNAME", os.environ.get("HOSTNAME", "desconhecido"))

def _usuario() -> str:
    return os.environ.get("USERNAME", os.environ.get("USER", "desconhecido"))


def adquirir_lock(lock_path: Path, timeout: int = LOCK_TIMEOUT) -> bool:
    """
    Adquire o lock compartilhado no OneDrive de forma segura.

    Fluxo:
      1. MELHORIA 1: jitter aleatório (0–3 s) para dessincronizar tentativas simultâneas.
      2. Se o lock existir, verifica se está morto (idade > LOCK_MAX_IDADE).
         Se morto, remove e tenta adquirir. Senão, aguarda e repete.
      3. Cria o lock atomicamente com os.O_EXCL.
      4. MELHORIA 2: aguarda LOCK_SYNC_WAIT (10 s) para o OneDrive sincronizar.
      5. MELHORIA 3: verificação dupla da propriedade (lê o lock 2 vezes com 1 s
         de intervalo) — garante que o sync estava completo na primeira leitura.
         Se outro PC sobrescreveu durante o wait, tenta novamente.
    """
    # MELHORIA 1: jitter — cada máquina espera um tempo aleatório diferente
    time.sleep(random.uniform(0, 3))

    inicio = time.time()

    while time.time() - inicio < timeout:

        # --- verifica lock existente ---
        if lock_path.exists():
            try:
                dados = json.loads(lock_path.read_text(encoding="utf-8"))
                criado_em = datetime.fromisoformat(dados.get("timestamp", ""))
                idade = (datetime.now() - criado_em).total_seconds()
                if idade > LOCK_MAX_IDADE:
                    lock_path.unlink(missing_ok=True)  # lock morto
                else:
                    time.sleep(2)
                    continue
            except Exception:
                lock_path.unlink(missing_ok=True)  # lock corrompido

        # --- tenta criar lock atomicamente ---
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump({
                    "maquina":   _maquina(),
                    "usuario":   _usuario(),
                    "timestamp": datetime.now().isoformat(),
                }, f)
        except FileExistsError:
            time.sleep(2)
            continue

        # MELHORIA 2: aguarda sync completo do OneDrive (10 s)
        time.sleep(LOCK_SYNC_WAIT)

        # MELHORIA 3: verificação dupla — lê o lock 2 vezes com 1 s de intervalo
        confirmacoes = 0
        for _ in range(2):
            try:
                dados = json.loads(lock_path.read_text(encoding="utf-8"))
                if dados.get("maquina") == _maquina() and dados.get("usuario") == _usuario():
                    confirmacoes += 1
            except Exception:
                pass
            time.sleep(1)

        if confirmacoes == 2:
            return True  # lock confirmado duas vezes — seguro prosseguir

        # outro PC sobrescreveu durante o wait — tenta novamente
        time.sleep(2)

    return False


def liberar_lock(lock_path: Path) -> None:
    try:
        lock_path.unlink(missing_ok=True)
    except Exception:
        pass


def _limpar_conflitos_onedrive(path_original: str) -> None:
    """
    MELHORIA 4: remove arquivos de conflito gerados pelo OneDrive após gravação.

    O OneDrive for Business cria cópias de conflito com nomes como:
      - "f_Diario_Bordo-NomeDaMaquina.xlsx"
      - "f_Diario_Bordo (NomeUsuario's conflicted copy YYYY-MM-DD).xlsx"
      - "f_Diario_Bordo (version conflict 1).xlsx"

    A heurística: qualquer arquivo na mesma pasta cujo nome comece com o stem
    do arquivo original, tenha a mesma extensão, mas nome diferente do original.
    """
    original = Path(path_original)
    pasta    = original.parent
    stem     = original.stem
    sufixo   = original.suffix

    try:
        for arquivo in pasta.iterdir():
            if arquivo.name == original.name:
                continue  # arquivo original — não remover
            if arquivo.suffix.lower() == sufixo.lower() and arquivo.stem.startswith(stem):
                try:
                    arquivo.unlink()
                except Exception:
                    pass
    except Exception:
        pass


# ==============================
# ESCRITA SEGURA NO ONEDRIVE — CORRIGIDA (Bug 1)
# ==============================

def _salvar_xlsx_seguro(df: pd.DataFrame, destino: str) -> None:
    """
    CORREÇÃO Bug 1: grava o DataFrame em um arquivo temporário FORA do OneDrive
    (em tempfile.gettempdir()) e depois copia para o destino com shutil.copy2.

    Isso evita que o OneDrive detecte um arquivo temporário .xlsx na pasta
    sincronizada e gere conflito de versão (cópia com nome da máquina).

    Parâmetros
    ----------
    df      : DataFrame já montado para gravação
    destino : caminho completo do arquivo .xlsx no OneDrive
    """
    tmp_dir  = Path(tempfile.gettempdir())
    pid      = os.getpid()
    nome_tmp = f"diario_bordo_tmp_{pid}_{int(time.time())}.xlsx"
    path_tmp = tmp_dir / nome_tmp

    ultimo_erro = None

    for tentativa in range(5):
        try:
            # 1. grava em local temporário (fora do OneDrive)
            df.to_excel(str(path_tmp), index=False)

            # 2. copia para o OneDrive substituindo o arquivo original
            shutil.copy2(str(path_tmp), destino)

            return  # sucesso

        except PermissionError as e:
            # arquivo aberto no Excel ou ainda sendo sincronizado
            ultimo_erro = e
            time.sleep(1.5)
        except Exception as e:
            ultimo_erro = e
            time.sleep(1)
        finally:
            if path_tmp.exists():
                try:
                    path_tmp.unlink()
                except Exception:
                    pass

    raise ultimo_erro if ultimo_erro else Exception(
        "Falha ao salvar após múltiplas tentativas."
    )


# ==============================
# CARREGAMENTO DAS BASES
# ==============================

@st.cache_data(ttl=30)
def carregar_controle():
    try:
        df = pd.read_excel(PATH_CONTROLE)
        if not df.empty:
            df.columns = df.columns.str.strip().str.upper()
        if "IDP_PROJETO" not in df.columns and "IDP" in df.columns:
            df["IDP_PROJETO"] = df["IDP"]
        return df
    except Exception as e:
        st.error(f"Erro ao carregar Controle de Projetos: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=30)
def carregar_diario():
    try:
        df = pd.read_excel(PATH_DIARIO)
        if not df.empty:
            df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")
        if "IDP_PROJETO" not in df.columns and "IDP" in df.columns:
            df["IDP_PROJETO"] = df["IDP"]
        return df
    except Exception as e:
        st.error(f"Erro ao carregar Diário de Bordo: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=30)
def carregar_apontamentos():
    try:
        df = pd.read_excel(PATH_APONT)
        if not df.empty:
            df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar Apontamentos: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=30)
def carregar_responsaveis():
    try:
        df = pd.read_excel(PATH_RESP)
        if not df.empty:
            df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar Responsáveis: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=30)
def carregar_tecnologia():
    try:
        df = pd.read_excel(PATH_TECNOLOGIA)
        if not df.empty:
            df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar Tecnologias: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=30)
def carregar_backlog():
    try:
        df = pd.read_excel(PATH_BACKLOG)
        if not df.empty:
            df.columns = df.columns.str.strip().str.upper()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar Backlog: {e}")
        return pd.DataFrame()


def salvar_controle(df: pd.DataFrame) -> bool:
    """Salva d_Controle_Projetos.xlsx usando escrita segura fora do OneDrive."""
    try:
        _salvar_xlsx_seguro(df, PATH_CONTROLE)
        _limpar_conflitos_onedrive(PATH_CONTROLE)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar Controle: {e}")
        return False


# ==============================
# PÁGINA: CADASTRAR / ATUALIZAR PROJETO
# ==============================

def page_cadastro_projeto():

    st.markdown("## Cadastro e Atualização de Projetos")
    st.caption("Cadastro e atualização centralizada da base 'Ficha de Projetos'")

    # ==============================
    # LIMPAR FORMULÁRIO
    # ==============================

    def limpar_formulario():
        chaves = [
            "idp_projeto_form", "cliente_form", "projeto_nome_form",
            "projeto_carimbo_form", "tecnologia_form", "status_sgp_form",
            "dv_form", "gcc_form", "gp_form", "gc_form", "csol_form",
            "cec_form", "coord_form", "data_assinatura_form",
            "data_implantacao_form", "vigencia_form", "qtd_pontos_form",
            "qtd_cctos_form", "cctos_cancelados_form", "valor_contrato_form",
            "receita_mensal_form", "risco_impl_form", "risco_oper_form",
            "objetos_form", "projeto_carregado",
        ]
        for chave in chaves:
            if chave in st.session_state:
                del st.session_state[chave]

    # ==============================
    # CARREGAR BASES
    # ==============================

    df_controle = carregar_controle()
    df_backlog  = carregar_backlog()
    df_resp     = carregar_responsaveis()
    df_tec      = carregar_tecnologia()

    if not df_controle.empty:
        df_controle.columns = df_controle.columns.str.strip().str.upper()

    # ==============================
    # HELPERS
    # ==============================

    def safe_value(valor):
        if pd.isna(valor):
            return "Selecione..."
        return str(valor).strip()

    def safe_int(valor):
        try:
            return int(float(valor))
        except Exception:
            return 0

    def safe_float(valor):
        try:
            if isinstance(valor, (int, float)):
                return float(valor)
            if pd.isna(valor):
                return 0.0
            valor = str(valor).strip()
            if "," in valor:
                valor = valor.replace(".", "").replace(",", ".")
            return float(valor)
        except Exception:
            return 0.0

    def parse_data_br(valor):
        try:
            return datetime.strptime(valor, "%d/%m/%Y")
        except Exception:
            return None

    def lista_unica(df, coluna):
        if coluna not in df.columns:
            return []
        return sorted(df[coluna].dropna().astype(str).str.strip().unique().tolist())

    # ==============================
    # PREENCHIMENTO AUTOMÁTICO
    # ==============================

    def preencher_formulario(linha):
        if linha is None:
            return

        def safe_date_str(valor):
            dt = pd.to_datetime(valor, errors="coerce")
            return "" if pd.isna(dt) else dt.strftime("%d/%m/%Y")

        st.session_state["idp_projeto_form"]       = str(linha.get("IDP_PROJETO", "") or "")
        st.session_state["cliente_form"]           = str(linha.get("CLIENTE", "") or "")
        st.session_state["projeto_nome_form"]      = str(linha.get("PROJETO", "") or "")
        st.session_state["projeto_carimbo_form"]   = str(linha.get("PROJETO_CARIMBO", "") or "")
        st.session_state["tecnologia_form"]        = str(linha.get("TECNOLOGIA", "Selecione...") or "Selecione...")
        st.session_state["status_sgp_form"]        = str(linha.get("STATUS_SGP", "Selecione...") or "Selecione...")
        st.session_state["status_geral_form"]      = str(linha.get("STATUS GERAL", "Selecione...") or "Selecione...")
        st.session_state["dv_form"]                = str(linha.get("DV", "") or "")
        st.session_state["gcc_form"]               = str(linha.get("GCC", "") or "")
        st.session_state["gp_form"]                = str(linha.get("GP", "") or "")
        st.session_state["gc_form"]                = str(linha.get("GC", "") or "")
        st.session_state["csol_form"]              = str(linha.get("CSOL", "") or "")
        st.session_state["cec_form"]               = str(linha.get("CEC", "") or "")
        st.session_state["coord_form"]             = str(linha.get("COORDENADOR", "") or "")
        st.session_state["data_assinatura_form"]   = safe_date_str(linha.get("DATA ASSINATURA CONTRATO"))
        st.session_state["data_implantacao_form"]  = safe_date_str(linha.get("DATA FINAL IMPLANTAÇÃO"))
        st.session_state["vigencia_form"]          = safe_int(linha.get("VIGÊNCIA CONTRATO (MESES)", 0))
        st.session_state["qtd_pontos_form"]        = safe_int(linha.get("QTD PONTOS (CIRCUITOS)", 0))
        st.session_state["qtd_cctos_form"]         = safe_int(linha.get("QTD CCTOS SOLICITADOS", 0))
        st.session_state["cctos_cancelados_form"]  = safe_int(linha.get("CCTOS CANCELADOS", 0))
        st.session_state["valor_contrato_form"]    = safe_float(linha.get("VALOR CONTRATO", 0))
        st.session_state["receita_mensal_form"]    = safe_float(linha.get("RECEITA MENSAL CONTRATUAL", 0))
        st.session_state["risco_impl_form"]        = safe_float(linha.get("RISCO MULTA IMPLANTAÇÃO", 0))
        st.session_state["risco_oper_form"]        = safe_float(linha.get("RISCO MULTA OPERAÇÃO", 0))
        st.session_state["objetos_form"]           = str(linha.get("OBJETOS", "") or "")

    # ==============================
    # MODO DE OPERAÇÃO
    # ==============================

    if "projeto_existente_sel" not in st.session_state:
        st.session_state["projeto_existente_sel"] = None

    st.divider()

    status_geral = "Selecione..."

    modo = st.radio(
        "Modo de operação",
        ["Novo Projeto", "Atualizar Projeto Existente"],
        key="modo_cadastro",
        horizontal=True
    )

    st.divider()

    projeto_contexto = None
    linha_projeto    = None

    if modo == "Atualizar Projeto Existente":

        if "PROJETO" not in df_controle.columns:
            st.error("A coluna PROJETO não foi encontrada em d_Controle_Projetos.")
            st.stop()

        lista_projetos = sorted(
            df_controle["PROJETO"].dropna().astype(str).str.strip().unique().tolist()
        )

        if not lista_projetos:
            st.warning("Nenhum projeto encontrado na base de controle.")
            st.stop()

        col_proj, col_status = st.columns([3, 2])

        with col_proj:
            projeto_contexto = st.selectbox(
                "Selecione o projeto para atualizar",
                ["Selecione..."] + lista_projetos,
                key="projeto_existente_sel"
            )

        if projeto_contexto != "Selecione...":
            df_proj = df_controle[
                df_controle["PROJETO"].astype(str).str.strip() == str(projeto_contexto).strip()
            ].copy()
            linha_projeto = df_proj.iloc[0] if not df_proj.empty else None

        if "projeto_carregado" not in st.session_state:
            st.session_state["projeto_carregado"] = None

        if projeto_contexto != st.session_state["projeto_carregado"]:
            st.session_state["projeto_carregado"] = projeto_contexto
            for key in list(st.session_state.keys()):
                if "_form" in key:
                    del st.session_state[key]
            if projeto_contexto != "Selecione..." and linha_projeto is not None:
                preencher_formulario(linha_projeto)
            st.rerun()

        if projeto_contexto != "Selecione..." and linha_projeto is not None:
            st.success(f"Projeto carregado para edição: {projeto_contexto}")

        with col_status:
            status_geral_default = "Selecione..."
            if linha_projeto is not None:
                status_geral_default = safe_value(linha_projeto.get("STATUS GERAL"))
            if "status_geral_form" not in st.session_state:
                st.session_state["status_geral_form"] = status_geral_default
            status_geral = st.selectbox(
                "STATUS GERAL",
                ["Selecione...", "Novo", "Em andamento", "Concluído"],
                key="status_geral_form"
            )

    # ==============================
    # CONTEXTO ATUAL
    # ==============================

    st.markdown("### Contexto atual")
    if modo == "Novo Projeto":
        st.write("Preparado para incluir um novo projeto.")
    else:
        if linha_projeto is not None:
            st.write(f"Projeto em edição: **{projeto_contexto}**")
        else:
            st.write("Nenhum projeto carregado.")

    # ==============================
    # IDENTIFICAÇÃO DO PROJETO
    # ==============================

    st.divider()
    st.markdown("### Identificação do Projeto")

    # aplicar sugestões do backlog
    if st.session_state.get("aplicar_backlog"):
        if st.session_state.get("cliente_sug"):
            st.session_state["cliente_form"] = str(st.session_state["cliente_sug"])
        if st.session_state.get("proj_carimbo_sug"):
            st.session_state["projeto_carimbo_form"] = str(st.session_state["proj_carimbo_sug"])
        st.session_state["aplicar_backlog"] = False

    # TECNOLOGIA
    if not df_tec.empty and "ATIVO" in df_tec.columns and "TECNOLOGIA" in df_tec.columns:
        lista_tecnologia = sorted(
            df_tec[df_tec["ATIVO"].astype(str).str.upper() == "SIM"]["TECNOLOGIA"]
            .dropna().astype(str).str.strip().unique().tolist()
        )
    else:
        lista_tecnologia = []

    opcoes_tecnologia = ["Selecione..."] + lista_tecnologia
    tec_default = ""
    if modo == "Atualizar Projeto Existente" and linha_projeto is not None:
        tec_default = str(linha_projeto.get("TECNOLOGIA", "")).strip()
    if tec_default not in opcoes_tecnologia:
        tec_default = "Selecione..."
    if "tecnologia_form" not in st.session_state:
        st.session_state["tecnologia_form"] = tec_default

    # STATUS SGP
    lista_status = ["Selecione...", "Planejamento", "Em implantação", "Ativação", "Operação Assistida", "Concluído"]
    status_default = ""
    if modo == "Atualizar Projeto Existente" and linha_projeto is not None:
        status_default = str(linha_projeto.get("STATUS_SGP", "")).strip()
    if st.session_state.get("status_sgp_sug"):
        status_default = str(st.session_state.get("status_sgp_sug")).strip()
    if status_default not in lista_status:
        status_default = "Selecione..."
    if "status_sgp_form" not in st.session_state:
        st.session_state["status_sgp_form"] = status_default

    col1, col2, col3, col4, col5, col6 = st.columns([1, 2, 2, 3, 2, 2])
    with col1:
        idp_projeto  = st.text_input("IDP_PROJETO",           key="idp_projeto_form")
    with col2:
        cliente      = st.text_input("CLIENTE",               key="cliente_form")
    with col3:
        projeto_nome = st.text_input("PROJETO",               key="projeto_nome_form")
    with col4:
        projeto_carimbo = st.text_input("PROJETO_CARIMBO",    key="projeto_carimbo_form")
    with col5:
        tecnologia   = st.selectbox("TECNOLOGIA",             opcoes_tecnologia, key="tecnologia_form")
    with col6:
        status_sgp   = st.selectbox("STATUS DO PROJETO (SGP)", lista_status,    key="status_sgp_form")

    # ==============================
    # RESPONSÁVEIS
    # ==============================

    lista_dv = lista_unica(df_controle, "DV")

    def lista_resp_por_funcao(funcao):
        if df_resp.empty or "FUNCAO" not in df_resp.columns or "ATIVO" not in df_resp.columns:
            return lista_unica(df_controle, funcao)
        mask = (
            df_resp["FUNCAO"].astype(str).str.strip().str.upper().str.contains(funcao, na=False)
        ) & (df_resp["ATIVO"].astype(str).str.strip().str.upper() == "SIM")
        return sorted(df_resp.loc[mask, "NOME"].dropna().astype(str).str.strip().unique().tolist())

    lista_gcc  = lista_resp_por_funcao("GCC")  or lista_unica(df_controle, "GCC")
    lista_gp   = lista_resp_por_funcao("GP")   or lista_unica(df_controle, "GP")
    lista_gc   = lista_resp_por_funcao("GC")   or lista_unica(df_controle, "GC")
    lista_csol = lista_resp_por_funcao("CSOL") or lista_unica(df_controle, "CSOL")
    lista_cec  = lista_resp_por_funcao("CEC")  or lista_unica(df_controle, "CEC")
    lista_coord= lista_resp_por_funcao("COORDENADOR") or lista_unica(df_controle, "COORDENADOR")
    opcoes_coord = ["Selecione..."] + lista_coord

    # construir mapa CEC→Coord a partir do controle + mapa manual
    mapa_cec_coord = {str(k).strip(): str(v).strip() for k, v in CEC_COORD_MAP.items()}
    if {"CEC", "COORDENADOR"}.issubset(df_controle.columns):
        pares = (
            df_controle[["CEC", "COORDENADOR"]]
            .dropna(subset=["CEC", "COORDENADOR"])
            .astype(str)
            .apply(lambda col: col.str.strip())
        )
        pares = pares[(pares["CEC"] != "") & (pares["COORDENADOR"] != "")]
        if not pares.empty:
            mapa_inferido = (
                pares.groupby("CEC")["COORDENADOR"]
                .agg(lambda s: s.iloc[0] if s.nunique() == 1 else None)
                .dropna().to_dict()
            )
            mapa_cec_coord = {**mapa_inferido, **mapa_cec_coord}

    def sincronizar_coord_por_cec():
        cec_sel = st.session_state.get("cec_form", "")
        coord_sugerido = mapa_cec_coord.get(cec_sel)
        if coord_sugerido and coord_sugerido in opcoes_coord:
            st.session_state["coord_form"] = coord_sugerido

    for campo, col_key in [
        ("dv_form", "DV"), ("gcc_form", "GCC"), ("gp_form", "GP"),
        ("gc_form", "GC"), ("csol_form", "CSOL"), ("cec_form", "CEC"),
        ("coord_form", "COORDENADOR"),
    ]:
        if campo not in st.session_state:
            default = ""
            if modo == "Atualizar Projeto Existente" and linha_projeto is not None:
                default = safe_value(linha_projeto.get(col_key))
            st.session_state[campo] = default

    st.divider()
    st.markdown("### Responsáveis do Projeto")

    col1, col2, col3, col4, col5, col6, col7 = st.columns([1.5, 3, 3, 3, 3, 3, 3])
    with col1:
        dv = st.selectbox("DV",                          ["Selecione..."] + lista_dv,   key="dv_form")
    with col2:
        gcc = st.selectbox("Gerente Coord. Cliente (GCC)", ["Selecione..."] + lista_gcc, key="gcc_form")
    with col3:
        gp = st.selectbox("Gerente de Projetos (GP)",    ["Selecione..."] + lista_gp,   key="gp_form")
    with col4:
        gc = st.selectbox("Gerente Comercial (GC)",      ["Selecione..."] + lista_gc,   key="gc_form")
    with col5:
        csol = st.selectbox("Consultor Soluções (CSOL)", ["Selecione..."] + lista_csol, key="csol_form")
    with col6:
        cec = st.selectbox("CEC",                        ["Selecione..."] + lista_cec,  key="cec_form",
                           on_change=sincronizar_coord_por_cec)
    with col7:
        coordenador = st.selectbox("Coordenador Técnico", opcoes_coord, key="coord_form")

    # ==============================
    # DADOS CONTRATUAIS
    # ==============================

    st.divider()
    st.markdown("### Dados Contratuais")

    def safe_date_str(valor):
        dt = pd.to_datetime(valor, errors="coerce")
        return "" if pd.isna(dt) else dt.strftime("%d/%m/%Y")

    for campo, getter, default in [
        ("data_assinatura_form",  lambda: safe_date_str(linha_projeto.get("DATA ASSINATURA CONTRATO")) if linha_projeto is not None else "", None),
        ("data_implantacao_form", lambda: safe_date_str(linha_projeto.get("DATA FINAL IMPLANTAÇÃO"))   if linha_projeto is not None else "", None),
        ("vigencia_form",         lambda: safe_int(linha_projeto.get("VIGÊNCIA CONTRATO (MESES)", 0))  if linha_projeto is not None else 0, 0),
    ]:
        if campo not in st.session_state:
            st.session_state[campo] = getter()

    col1, col2, col3 = st.columns(3)
    with col1:
        data_assinatura_str = st.text_input("Data Assinatura Contrato (dd/mm/aaaa)", key="data_assinatura_form")
        data_assinatura = parse_data_br(data_assinatura_str)
        if data_assinatura_str and not data_assinatura:
            st.warning("Formato inválido. Use dd/mm/aaaa")
    with col2:
        data_implantacao_str = st.text_input("Data Final Implantação (dd/mm/aaaa)", key="data_implantacao_form")
        data_final_implantacao = parse_data_br(data_implantacao_str)
        if data_implantacao_str and not data_final_implantacao:
            st.warning("Formato inválido. Use dd/mm/aaaa")
    with col3:
        vigencia_contrato = st.number_input("Vigência Contrato (meses)", min_value=0, step=1, key="vigencia_form")

    st.markdown("#### Quantidades")
    for campo, getter in [
        ("qtd_pontos_form",      lambda: safe_int(linha_projeto.get("QTD PONTOS (CIRCUITOS)", 0))  if linha_projeto is not None else 0),
        ("qtd_cctos_form",       lambda: safe_int(linha_projeto.get("QTD CCTOS SOLICITADOS", 0))   if linha_projeto is not None else 0),
        ("cctos_cancelados_form",lambda: safe_int(linha_projeto.get("CCTOS CANCELADOS", 0))        if linha_projeto is not None else 0),
    ]:
        if campo not in st.session_state:
            st.session_state[campo] = getter()

    col1, col2, col3 = st.columns(3)
    with col1:
        qtd_pontos      = st.number_input("Qtd Pontos (Previsão inicial)", min_value=0, step=1,   key="qtd_pontos_form")
    with col2:
        qtd_cctos       = st.number_input("Qtd Cctos Solicitados",         min_value=0, step=1,   key="qtd_cctos_form")
    with col3:
        cctos_cancelados= st.number_input("Cctos Cancelados",             min_value=0, step=1,   key="cctos_cancelados_form")

    st.markdown("#### Valores Financeiros")
    for campo, getter in [
        ("valor_contrato_form",  lambda: safe_float(linha_projeto.get("VALOR CONTRATO", 0))            if linha_projeto is not None else 0.0),
        ("receita_mensal_form",  lambda: safe_float(linha_projeto.get("RECEITA MENSAL CONTRATUAL", 0)) if linha_projeto is not None else 0.0),
        ("risco_impl_form",      lambda: safe_float(linha_projeto.get("RISCO MULTA IMPLANTAÇÃO", 0))   if linha_projeto is not None else 0.0),
        ("risco_oper_form",      lambda: safe_float(linha_projeto.get("RISCO MULTA OPERAÇÃO", 0))      if linha_projeto is not None else 0.0),
    ]:
        if campo not in st.session_state:
            st.session_state[campo] = getter()

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        valor_contrato  = st.number_input("Valor Contrato (R$)",             min_value=0.0, step=1000.0, format="%.2f", key="valor_contrato_form")
    with col2:
        receita_mensal  = st.number_input("Receita Mensal Contratual (R$)",  min_value=0.0, step=100.0,  format="%.2f", key="receita_mensal_form")
    with col3:
        risco_multa_impl= st.number_input("Risco Multa Implantação (%)",     min_value=0.0, step=100.0,  format="%.2f", key="risco_impl_form")
    with col4:
        risco_multa_oper= st.number_input("Risco Multa Operação (%)",        min_value=0.0, step=100.0,  format="%.2f", key="risco_oper_form")

    st.markdown("#### Objeto do Projeto")
    if "objetos_form" not in st.session_state:
        st.session_state["objetos_form"] = str(linha_projeto.get("OBJETOS", "") or "") if linha_projeto is not None else ""
    objetos = st.text_area("Descrição / Objeto do Projeto", height=120, key="objetos_form")

    # ==============================
    # SALVAR PROJETO
    # ==============================

    st.divider()
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("🧹 Limpar", use_container_width=True):
            limpar_formulario()
            if "projeto_existente_sel" in st.session_state:
                del st.session_state["projeto_existente_sel"]
            st.rerun()
    with col_btn2:
        salvar_projeto = st.button("💾 Salvar Projeto", use_container_width=True, type="primary")

    if salvar_projeto:
        if not idp_projeto or not cliente or not projeto_nome:
            st.warning("Preencha pelo menos IDP_PROJETO, CLIENTE e PROJETO.")
            st.stop()

        if modo == "Novo Projeto":
            existe = df_controle[df_controle["IDP_PROJETO"].astype(str).str.strip() == str(idp_projeto).strip()]
            if not existe.empty:
                st.error("Já existe um projeto com esse IDP_PROJETO.")
                st.stop()

        try:
            df_controle_atual = carregar_controle().copy()
            df_controle_atual.columns = df_controle_atual.columns.str.strip().str.upper()

            status_geral_final = "Novo" if modo == "Novo Projeto" else ("" if status_geral == "Selecione..." else status_geral)

            registro = {
                "IDP_PROJETO": idp_projeto, "CLIENTE": cliente, "PROJETO": projeto_nome,
                "PROJETO_CARIMBO": projeto_carimbo,
                "CARIMBO_PREFIXO": st.session_state.get("carimbo_sug", ""),
                "TECNOLOGIA": tecnologia, "DV": dv, "GCC": gcc, "GP": gp,
                "GC": gc, "CSOL": csol, "CEC": cec, "COORDENADOR": coordenador,
                "STATUS_SGP": status_sgp, "STATUS GERAL": status_geral_final,
                "DATA ASSINATURA CONTRATO": data_assinatura,
                "DATA FINAL IMPLANTAÇÃO": data_final_implantacao,
                "VIGÊNCIA CONTRATO (MESES)": vigencia_contrato,
                "QTD PONTOS (CIRCUITOS)": qtd_pontos,
                "QTD CCTOS SOLICITADOS": qtd_cctos, "CCTOS CANCELADOS": cctos_cancelados,
                "VALOR CONTRATO": valor_contrato, "RECEITA MENSAL CONTRATUAL": receita_mensal,
                "RISCO MULTA IMPLANTAÇÃO": risco_multa_impl, "RISCO MULTA OPERAÇÃO": risco_multa_oper,
                "OBJETOS": objetos,
            }

            filtro = df_controle_atual["IDP_PROJETO"].astype(str).str.strip() == str(idp_projeto).strip()
            if filtro.any():
                df_controle_atual.loc[filtro, list(registro.keys())] = list(registro.values())
            else:
                df_controle_atual = pd.concat([df_controle_atual, pd.DataFrame([registro])], ignore_index=True)

            salvo = salvar_controle(df_controle_atual)
            if not salvo:
                st.error("Erro ao salvar o projeto.")
                st.stop()

            st.success("✅ Projeto salvo com sucesso!")
            st.cache_data.clear()
            st.rerun()

        except Exception as e:
            st.error(f"Erro ao salvar projeto: {e}")

    # ==============================
    # BUSCA AUTOMÁTICA NO BACKLOG
    # ==============================

    if idp_projeto and "IDP_PROJETO" in df_backlog.columns:
        df_bk = df_backlog[df_backlog["IDP_PROJETO"].astype(str).str.strip() == str(idp_projeto).strip()]
        if not df_bk.empty:
            dados_bk = df_bk.iloc[0]
            if st.session_state.get("idp_processado") != str(idp_projeto).strip():
                st.session_state["cliente_sug"]       = str(dados_bk.get("CLIENTE", "") or "")
                st.session_state["proj_carimbo_sug"]  = str(dados_bk.get("CARIMBO_PROJETO", "") or "")
                st.session_state["status_sgp_sug"]    = str(dados_bk.get("STATUS_PROJETO", "") or "")
                st.session_state["carimbo_sug"]       = str(dados_bk.get("CARIMBO_PREFIXO", "") or "")
                st.session_state["aplicar_backlog"]   = True
                st.session_state["idp_processado"]    = str(idp_projeto).strip()
                st.rerun()
            st.success("Projeto encontrado automaticamente no backlog.")


# ==============================
# PÁGINA: ATUALIZAR DIÁRIO DE BORDO
# ==============================

def page_atualizar_diario():

    st.markdown("## Atualizar Diário de Bordo")

    st.markdown("""
    <style>
    div.stButton > button:first-child {
        background-color: #1E293B; color: white;
        border-radius: 8px; border: none; height: 45px;
    }
    div.stButton > button:first-child:hover,
    div.stButton > button:first-child:active,
    div.stButton > button:first-child:focus {
        background-color: #3B82F6; color: white; box-shadow: none;
    }
    </style>
    """, unsafe_allow_html=True)

    col_esq, col_centro, col_dir = st.columns([1, 2, 1])

    with col_centro:

        # SESSION STATE INICIAL
        for key, default in [
            ("limpar_form", False), ("area_form", "Selecione..."),
            ("status_form", "Selecione..."), ("data_realizado_form", date.today()),
            ("obs_form", ""), ("msg_sucesso", False), ("pendencia_form", "Não"),
        ]:
            if key not in st.session_state:
                st.session_state[key] = default

        if st.session_state["limpar_form"]:
            st.session_state["area_form"]            = "Selecione..."
            st.session_state["status_form"]          = "Selecione..."
            st.session_state["data_realizado_form"]  = date.today()
            st.session_state["obs_form"]             = ""
            st.session_state["pendencia_form"]       = "Não"
            st.session_state["limpar_form"]          = False

        if st.session_state["msg_sucesso"]:
            st.success("✅ Registro salvo com sucesso!")
            st.session_state["msg_sucesso"] = False

        # CARREGAR BASES
        df_controle   = carregar_controle()
        df_diario     = carregar_diario()
        df_apontamentos = carregar_apontamentos()

        if not df_controle.empty:
            df_controle.columns   = df_controle.columns.str.strip().str.upper()
        if not df_diario.empty:
            df_diario.columns     = df_diario.columns.str.strip().str.upper()
        if not df_apontamentos.empty:
            df_apontamentos.columns = df_apontamentos.columns.str.strip().str.upper()

        # SELEÇÃO DO PROJETO
        st.divider()
        if "PROJETO" not in df_controle.columns:
            st.error("Coluna PROJETO não encontrada na base de controle. Verifique o arquivo d_Controle_Projetos.xlsx.")
            st.stop()
        lista_projetos = sorted(df_controle["PROJETO"].dropna().astype(str).unique())
        projeto = st.selectbox("Projeto", lista_projetos)

        st.divider()
        st.markdown(f"### 📋 {projeto}")

        # STATUS MACRO
        st.divider()
        lista_status_macro = [
            "Selecione...", "Aprovação CAPEX", "Aquisição Equipamentos PJE",
            "Assinatura contrato", "Cadastro", "Em implantação", "Finalizado",
            "Finalizado com Pendência", "Homologação Solução", "Projeto",
        ]

        status_macro_atual = ""
        if projeto and "STATUS MACRO" in df_controle.columns:
            linha_ctrl = df_controle[df_controle["PROJETO"].astype(str).str.strip() == str(projeto).strip()]
            if not linha_ctrl.empty:
                status_macro_atual = str(linha_ctrl.iloc[0].get("STATUS MACRO", "")).strip()

        status_macro_novo = st.selectbox(
            "STATUS MACRO",
            lista_status_macro,
            index=lista_status_macro.index(status_macro_atual) if status_macro_atual in lista_status_macro else 0
        )

        if st.button("Atualizar Status Macro"):
            if status_macro_novo == "Selecione...":
                st.warning("Selecione um status válido.")
            else:
                try:
                    df_ctrl = carregar_controle().copy()
                    filtro = df_ctrl["PROJETO"].astype(str).str.strip() == str(projeto).strip()
                    if not filtro.any():
                        st.error("Projeto não encontrado na base de controle.")
                    else:
                        df_ctrl.loc[filtro, "STATUS MACRO"] = status_macro_novo
                        if salvar_controle(df_ctrl):
                            st.success("Status Macro atualizado com sucesso!")
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error("Erro ao salvar a base de controle.")
                except Exception as e:
                    st.error(f"Erro ao atualizar: {e}")

        # DADOS DO PROJETO
        df_proj = df_controle[df_controle["PROJETO"].astype(str) == str(projeto)]
        if df_proj.empty:
            st.error("Projeto não encontrado na base de controle.")
            st.stop()

        linha_proj = df_proj.iloc[0]
        projeto_carimbo = None
        if "PROJETO_CARIMBO" in df_proj.columns:
            projeto_carimbo = str(linha_proj["PROJETO_CARIMBO"]).strip().split(" - ")[0].strip()

        idp_projeto = linha_proj.get("IDP_PROJETO")

        cec_projeto = None
        if "CEC - PROJETO" in df_diario.columns:
            vals = df_diario.loc[df_diario["PROJETO"].astype(str) == str(projeto), "CEC - PROJETO"].dropna().astype(str).unique()
            if len(vals) > 0:
                cec_projeto = vals[0]
        elif "CEC" in df_controle.columns:
            cec_projeto = linha_proj.get("CEC")

        # ETAPA / APONTAMENTO
        if "STATUS_MACRO" not in df_apontamentos.columns:
            st.error("Coluna STATUS_MACRO não encontrada em d_apontamentos. Verifique o arquivo.")
            st.stop()
        lista_macros = sorted(df_apontamentos["STATUS_MACRO"].dropna().astype(str).unique())
        macro = st.selectbox("Etapa", lista_macros)

        df_macro = df_apontamentos[df_apontamentos["STATUS_MACRO"].astype(str) == str(macro)].copy()
        lista_apontamentos = sorted(df_macro["APONTAMENTOS_PADRÃO"].dropna().astype(str).unique())
        apontamento = st.selectbox("Apontamento", lista_apontamentos)

        pendencia = st.selectbox("Possui Pendência?", ["Não", "Sim"], key="pendencia_form")

        df_ap = df_macro[df_macro["APONTAMENTOS_PADRÃO"].astype(str) == str(apontamento)]
        if df_ap.empty:
            st.error("Apontamento não encontrado na dimensão d_apontamentos.")
            st.stop()
        apontamento_sk = df_ap.iloc[0].get("APONTAMENTO_SK") if "APONTAMENTO_SK" in df_ap.columns else None

        # ÁREA
        lista_areas = sorted(df_diario["RESPONSAVEL_AREA"].dropna().astype(str).unique()) if "RESPONSAVEL_AREA" in df_diario.columns else []
        area = st.selectbox("Área", ["Selecione..."] + lista_areas, key="area_form")

        # STATUS BOTÕES
        st.divider()
        st.markdown("**Status Apontamento**")
        col1, col2, col3 = st.columns(3, gap="small")
        with col1:
            if st.button("⚪ Não iniciado", use_container_width=True):
                st.session_state["status_form"] = "Não iniciado"
        with col2:
            if st.button("🟡 Em andamento", use_container_width=True):
                st.session_state["status_form"] = "Em andamento"
        with col3:
            if st.button("🔵 Concluído", use_container_width=True):
                st.session_state["status_form"] = "Concluído"

        status = st.session_state["status_form"]
        st.caption(f"Status selecionado: {status}")

        st.divider()

        data_realizado = st.date_input("Data em que o evento aconteceu", key="data_realizado_form")
        data_obs = date.today()
        observacao = st.text_area("Observação", height=120, key="obs_form")

        # ==============================
        # SALVAR APONTAMENTO — CORRIGIDO (Bug 1 + Bug 2)
        # ==============================

        salvar = st.button("💾 Salvar Apontamento", use_container_width=True, type="primary")

        if salvar:
            # adquirir_lock agora usa LOCK_DIARIO (no OneDrive) com sync wait
            if not adquirir_lock(LOCK_DIARIO):
                st.error("❌ Outro usuário está gravando no Diário neste momento. Tente novamente em alguns segundos.")
                st.stop()

            if area == "Selecione..." or status == "Selecione...":
                liberar_lock(LOCK_DIARIO)
                st.warning("Selecione Área e Status antes de salvar.")
                st.stop()

            try:
                # Leitura fresh do arquivo (sempre fora do cache)
                df_existente = pd.read_excel(PATH_DIARIO)
                df_existente.columns = df_existente.columns.str.strip().str.upper()

                novo_registro = {
                    "PROJETO": projeto, "PROJETO_CARIMBO": projeto_carimbo,
                    "IDP_PROJETO": idp_projeto, "APONTAMENTOS": apontamento,
                    "STATUS_MACRO": macro, "APONTAMENTO_ATIVO": apontamento,
                    "PENDENCIA_APONTAMENTO": pendencia, "RESPONSAVEL_AREA": area,
                    "RESPONSAVEL_NOME": None, "CEC - PROJETO": cec_projeto,
                    "DATA_REALIZADO": pd.Timestamp(data_realizado),
                    "STATUS": status, "DATA_OBS": pd.Timestamp(data_obs),
                    "OBSERVACAO": observacao, "APONTAMENTO_LEGADO_ORIGEM": None,
                    "APONTAMENTO_SK": apontamento_sk,
                    "DATA_REGISTRO": pd.Timestamp.now(),
                }

                df_novo  = pd.DataFrame([novo_registro]).reindex(columns=df_existente.columns)
                df_final = pd.concat([df_existente, df_novo], ignore_index=True)

                # CORREÇÃO Bug 1: usa _salvar_xlsx_seguro (temp fora do OneDrive)
                _salvar_xlsx_seguro(df_final, PATH_DIARIO)

                # MELHORIA 4: remove conflitos gerados pelo OneDrive após a gravação
                _limpar_conflitos_onedrive(PATH_DIARIO)

                # VERIFICAÇÃO PÓS-GRAVAÇÃO
                df_check = pd.read_excel(PATH_DIARIO)
                df_check.columns = df_check.columns.str.strip().str.upper()
                filtro_check = (
                    (df_check["PROJETO"].astype(str) == str(projeto)) &
                    (df_check["APONTAMENTO_ATIVO"].astype(str) == str(apontamento)) &
                    (pd.to_datetime(df_check["DATA_OBS"], errors="coerce") == pd.Timestamp(data_obs))
                )
                if not filtro_check.any():
                    raise Exception("Registro não encontrado após gravação no arquivo.")

                st.session_state["msg_sucesso"] = True
                st.session_state["limpar_form"] = True
                st.cache_data.clear()
                st.rerun()

            except PermissionError:
                st.error("❌ O arquivo f_Diario_Bordo.xlsx está aberto no Excel. Feche e tente novamente.")
            except Exception as e:
                st.error(f"❌ Erro ao salvar: {e}")
            finally:
                liberar_lock(LOCK_DIARIO)   # SEMPRE libera, mesmo em erro

        # HISTÓRICO DO PROJETO
        st.divider()
        st.markdown("### 📜 Histórico do Projeto")

        df_hist = df_diario[df_diario["PROJETO"].astype(str).str.strip() == str(projeto)].copy()

        if df_hist.empty:
            st.info("Nenhum apontamento registrado para este projeto.")
        else:
            df_hist["DATA_OBS"]      = pd.to_datetime(df_hist["DATA_OBS"],      errors="coerce")
            df_hist["DATA_REGISTRO"] = pd.to_datetime(df_hist["DATA_REGISTRO"], errors="coerce")
            df_hist = df_hist.sort_values("DATA_REGISTRO", ascending=False, na_position="last").reset_index(drop=True)

            mapa_cores = {"Concluído": "#60A5FA", "Em andamento": "#FACC15", "Não iniciado": "#64748B"}
            html_itens = ""

            for _, row in df_hist.iterrows():
                data     = row.get("DATA_REALIZADO") or row.get("DATA_OBS")
                data_txt = data.strftime("%d/%m/%Y") if pd.notnull(data) else "-"
                macro_txt    = str(row.get("STATUS_MACRO", "")).strip()
                apont_txt    = str(row.get("APONTAMENTO_ATIVO", "")).strip()
                status_txt   = str(row.get("STATUS", "")).strip()
                obs_txt      = html_mod.escape(str(row.get("OBSERVACAO", "")).strip() or "Sem observação")
                area_txt     = html_mod.escape(str(row.get("RESPONSAVEL_AREA", "")).strip())
                pendencia_txt= str(row.get("PENDENCIA_APONTAMENTO", "")).strip()
                cor_status   = "#EF4444" if pendencia_txt == "Sim" else mapa_cores.get(status_txt, "#94A3B8")

                html_itens += f"""
                <div style="border-left:4px solid {cor_status};padding:12px 15px;margin-bottom:12px;
                     background-color:#1E293B;border-radius:10px;box-shadow:0 4px 14px rgba(0,0,0,.4);">
                  <div style="font-size:12px;color:#94A3B8;">{data_txt} • {macro_txt}</div>
                  <div style="font-size:16px;font-weight:600;color:#F59E0B;margin-top:4px;">
                    {"⚠️ " if pendencia_txt=="Sim" else ""}{apont_txt}</div>
                  <div style="font-size:13px;color:#E2E8F0;margin-top:4px;">
                    Status: <b style="color:{cor_status};">{status_txt}</b></div>
                  <div style="font-size:13px;color:#CBD5E1;margin-top:8px;">{obs_txt}</div>
                  <div style="font-size:11px;color:#94A3B8;margin-top:8px;">Área: {area_txt}</div>
                </div>
                """

            components.html(
                f'<div style="font-family:Segoe UI,Arial,sans-serif;">{html_itens}</div>',
                height=520, scrolling=True
            )


# ==============================
# APP PRINCIPAL
# ==============================

st.set_page_config(
    page_title="Diário de Bordo — Entrada de Dados",
    layout="wide",
    page_icon="📋"
)

st.markdown("""
<style>
.block-container { padding-left:1.5rem; padding-right:1.5rem; padding-top:1rem; }
</style>
""", unsafe_allow_html=True)

pagina = st.sidebar.radio(
    "Navegação",
    ["📁 Cadastrar / Atualizar Projeto", "📝 Atualizar Diário de Bordo"],
)

st.sidebar.markdown("---")
st.sidebar.caption(f"Usuário: {os.getenv('USERNAME', 'desconhecido')}")
st.sidebar.caption(f"OneDrive: {_ONEDRIVE}")

if pagina == "📁 Cadastrar / Atualizar Projeto":
    page_cadastro_projeto()
else:
    page_atualizar_diario()
