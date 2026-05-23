"""
=======================================================
 BABE KOMBUCHA — Plataforma de Visitas ao Varejo
=======================================================
 Funcionalidades:
   1. Importar dados mensais de vendas
   2. Top 55 varejistas + Registro de visitas por dia
   3. Planos de Ação — controle de encaminhamentos
=======================================================
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
from datetime import datetime, date
import os
import json
import numpy as np
import io

# ─────────────────────────────────────────
#  CONFIGURAÇÃO GERAL
# ─────────────────────────────────────────
st.set_page_config(
    page_title="Babe Kombucha – Visitas Varejo",
    page_icon="🍵",
    layout="wide",
)

APP_DIR      = os.path.dirname(os.path.abspath(__file__))
VISITAS_FILE = os.path.join(APP_DIR, "visitas.json")   # histórico de visitas por mês
ROUTESCAPE_SHEET = "routescape"

COL_ZONE        = 2
COL_ACCOUNT     = 5
COL_ADDRESS     = 6
COL_CITY        = 7
COL_ZIPCODE     = 8
COL_DATES_START = 9

MESES_PT = {
    1:"Janeiro", 2:"Fevereiro", 3:"Março",    4:"Abril",
    5:"Maio",    6:"Junho",     7:"Julho",     8:"Agosto",
    9:"Setembro",10:"Outubro",  11:"Novembro", 12:"Dezembro"
}

GEO_CACHE_FILE = os.path.join(APP_DIR, "geo_cache_v3.json")

# ─────────────────────────────────────────────────────────────
#  FUNÇÕES — AGRUPAMENTO GEOGRÁFICO
#  Algoritmo: RAIO A PARTIR DO VAREJISTA MAIS AO NORTE
#  1. Dentre os varejistas ainda não agendados, escolhe o
#     mais ao NORTE (maior latitude) — este é o "âncora do dia"
#  2. Calcula a distância em linha reta de todos os outros
#     não-agendados até esse âncora
#  3. Pega os (max_por_dia - 1) mais próximos em raio — eles
#     formam o DIA junto com o âncora
#  4. Remove esses varejistas da lista e recomeça o DIA SEGUINTE
#     escolhendo o mais ao norte dos que sobraram
#  Resultado: cada dia é um cluster compacto em torno de um
#  ponto de referência ao norte — sem saltos geográficos
# ─────────────────────────────────────────────────────────────

_ZIP_COORDS = {
    92037:(32.8328,-117.2713), 92130:(32.9540,-117.2282), 92122:(32.8701,-117.2112),
    92109:(32.7940,-117.2521), 92107:(32.7479,-117.2497), 92106:(32.7260,-117.2437),
    92101:(32.7157,-117.1611), 92103:(32.7492,-117.1677), 92104:(32.7398,-117.1305),
    92116:(32.7571,-117.1127), 92105:(32.7315,-117.0885), 92117:(32.8145,-117.1973),
    92111:(32.8078,-117.1513), 92108:(32.7730,-117.1445), 92102:(32.7206,-117.1214),
    92114:(32.6991,-117.0477), 92115:(32.7407,-117.0764), 92014:(32.9595,-117.2653),
    92075:(32.9909,-117.2712), 92024:(33.0369,-117.2920), 92008:(33.1581,-117.3506),
    92009:(33.1350,-117.3115), 92010:(33.1641,-117.3027), 92011:(33.0922,-117.3100),
    92054:(33.1959,-117.3795), 92056:(33.1892,-117.3063), 92057:(33.2210,-117.3558),
    92058:(33.2089,-117.3790), 92025:(33.1192,-117.0864), 92026:(33.1469,-117.1032),
    92027:(33.1299,-117.0499), 92029:(33.0913,-117.0739), 92069:(33.1434,-117.1660),
    92078:(33.1558,-117.1948), 92083:(33.2000,-117.2425), 92084:(33.2090,-117.2082),
    92085:(33.1800,-117.2200), 92064:(32.9628,-117.0359), 92067:(33.0136,-117.1992),
    92127:(33.0284,-117.1149), 92128:(33.0128,-117.0812), 92129:(32.9701,-117.1163),
    92131:(32.9334,-117.0758), 92126:(32.9123,-117.1427), 92121:(32.8985,-117.1882),
    92071:(32.8384,-116.9739), 92020:(32.7948,-116.9625), 92021:(32.8024,-116.9261),
    92019:(32.7480,-116.9227), 92040:(32.8665,-116.9326), 92118:(32.6860,-117.1831),
    91910:(32.6401,-117.0842), 91911:(32.6038,-117.0600), 91950:(32.6781,-117.0992),
    91945:(32.7147,-117.0333), 92120:(32.7899,-117.0730), 91942:(32.7880,-117.0243),
    91941:(32.7648,-116.9877), 91977:(32.7350,-117.0022), 92139:(32.6770,-117.0322),
    92154:(32.5680,-117.0641), 92055:(33.2000,-117.4000), 92028:(33.3734,-117.1506),
    92003:(33.3100,-117.1700), 92065:(33.0285,-116.9697), 92060:(33.2900,-116.8500),
    92070:(33.1200,-116.7500), 92093:(32.8800,-117.2340), 92007:(33.0100,-117.2790),
    92110:(32.7550,-117.2040), 92119:(32.8020,-117.0060), 91914:(32.6500,-117.0000),
    91932:(32.5780,-117.1330),
}


def _load_geo_cache() -> dict:
    if os.path.exists(GEO_CACHE_FILE):
        with open(GEO_CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_geo_cache(cache: dict):
    with open(GEO_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2)


def _geocode(address: str, city: str, zipcode, cache: dict):
    """
    Converte endereço em coordenadas lat/lng.
    Usa ZIP como base + hash do endereço para micro-deslocamento
    determinístico (garante que o mesmo endereço sempre caia no
    mesmo ponto, diferenciando lojas na mesma rua).
    """
    key = f"{address}|{city}|{zipcode}"
    if key in cache and cache[key]:
        return tuple(cache[key])
    base_lat, base_lng = 32.7157, -117.1611   # fallback: centro de San Diego
    try:
        z = int(str(zipcode).strip())
        if z in _ZIP_COORDS:
            base_lat, base_lng = _ZIP_COORDS[z]
    except Exception:
        pass
    # Micro-deslocamento baseado no hash do endereço completo (≈300 m max)
    h    = hash(str(address)) & 0xFFFFFFFF
    dlat = ((h & 0xFFFF) / 0xFFFF - 0.5) * 0.003
    dlng = (((h >> 16) & 0xFFFF) / 0xFFFF - 0.5) * 0.003
    coords = (base_lat + dlat, base_lng + dlng)
    cache[key] = list(coords)
    return coords


def agrupar_por_proximidade(df: pd.DataFrame, max_por_dia: int) -> pd.DataFrame:
    """
    Agrupamento por RAIO a partir do varejista mais ao norte.

    Lógica exata (conforme especificação do cliente):
      1. Geocodifica TODOS os varejistas pelo endereço completo
         (rua + número + cidade + ZIP) — cada endereço vira um
         ponto lat/lng no mapa
      2. Enquanto houver varejistas sem dia atribuído:
           a) Âncora do dia = o MAIS AO NORTE (maior latitude)
              entre os que ainda não foram agendados
           b) Calcula distância geodésica do âncora até CADA
              um dos outros ainda-não-agendados
           c) Ordena por distância crescente e pega os
              (max_por_dia - 1) mais próximos
           d) Âncora + esses N-1 = visitas do DIA atual
           e) Marca todos como agendados e incrementa dia
      3. No último dia pode sobrar menos que max_por_dia

    Observação: o "mais ao norte" recomeça a cada dia, então
    o dia 2 pode começar numa região bem distinta do dia 1 —
    isso é intencional para evitar deslocamentos enormes num
    mesmo dia.
    """
    cache = _load_geo_cache()

    lats, lngs = [], []
    for _, row in df.iterrows():
        lat, lng = _geocode(row["Endereço"], row["Cidade"], row["ZIP"], cache)
        lats.append(lat)
        lngs.append(lng)

    _save_geo_cache(cache)

    coords = np.array(list(zip(lats, lngs)))
    n      = len(coords)
    dias   = [0] * n          # 0 = ainda não agendado

    dia_atual = 0
    while True:
        nao_agendados = [i for i in range(n) if dias[i] == 0]
        if not nao_agendados:
            break

        dia_atual += 1

        # (a) Âncora = mais ao norte dos que sobraram
        ancora = max(nao_agendados, key=lambda i: coords[i][0])
        ax, ay = coords[ancora]

        # (b,c) Ordena os outros pela distância ao âncora
        outros = [i for i in nao_agendados if i != ancora]
        outros.sort(
            key=lambda i: (coords[i][0] - ax) ** 2 + (coords[i][1] - ay) ** 2
        )

        # (d) Pega os (max_por_dia - 1) mais próximos
        selecionados = [ancora] + outros[: max_por_dia - 1]

        # (e) Marca todos como pertencentes ao dia atual
        for idx in selecionados:
            dias[idx] = dia_atual

    df = df.copy()
    df["Dia"] = dias
    return df

# ─────────────────────────────────────────
#  FUNÇÕES — VISITAS (JSON)
#  Cada mês é uma "fatia" independente.
#  Meses passados ficam congelados.
# ─────────────────────────────────────────

@st.cache_data(show_spinner=False)
def _load_visitas_cached(mtime: float) -> dict:
    """Cache da leitura do visitas.json — invalida quando o mtime muda."""
    if os.path.exists(VISITAS_FILE):
        with open(VISITAS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def load_visitas() -> dict:
    """Carrega o arquivo visitas.json com cache baseado em mtime."""
    mtime = os.path.getmtime(VISITAS_FILE) if os.path.exists(VISITAS_FILE) else 0.0
    return _load_visitas_cached(mtime)


def save_visitas(data: dict):
    """Salva o dicionário de visitas no arquivo JSON e limpa o cache."""
    with open(VISITAS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    _load_visitas_cached.clear()


def mes_key(mes: int, ano: int) -> str:
    """Ex: (1, 2026) → '2026-01'"""
    return f"{ano}-{mes:02d}"


def visita_key(account: str, address: str) -> str:
    """Chave única por varejista dentro do mês."""
    return f"{str(account).strip().upper()}|{str(address).strip().upper()}"


def get_visita(mes: int, ano: int, account: str, address: str) -> dict:
    visitas = load_visitas()
    return visitas.get(mes_key(mes, ano), {}).get(visita_key(account, address), {})


def salvar_visita(mes: int, ano: int, account: str, address: str, dados: dict):
    visitas = load_visitas()
    mk = mes_key(mes, ano)
    vk = visita_key(account, address)
    if mk not in visitas:
        visitas[mk] = {}
    visitas[mk][vk] = dados
    save_visitas(visitas)


# ─────────────────────────────────────────
#  FUNÇÕES — EXCEL (routescape)
# ─────────────────────────────────────────

def encontrar_arquivo_historico():
    # Ignora arquivos temporários e gerados pelo próprio app
    _ignorar = {"_temp_mensal.xlsx"}
    arquivos = [
        f for f in os.listdir(APP_DIR)
        if f.endswith(".xlsx")
        and f not in _ignorar
        and not f.startswith("Planos_Acao_")
        and not f.startswith("~$")            # arquivo aberto no Excel
    ]
    lenore = [f for f in arquivos if "lenore" in f.lower() or "john" in f.lower()]
    return lenore if lenore else arquivos


@st.cache_data(show_spinner=False)
def _carregar_routescape_dados(filepath: str, mtime: float):
    """
    Lê SOMENTE os dados do routescape (sem o objeto workbook).
    Esse resultado é cacheado. Usado por toda leitura read-only
    (Top 55, sidebar, ler_dados_mes_importado, etc.)
    """
    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb[ROUTESCAPE_SHEET]
    headers, rows, date_cols = [], [], []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = list(row)
            for j, h in enumerate(headers):
                if j >= COL_DATES_START and isinstance(h, datetime):
                    date_cols.append((j, h))
        else:
            rows.append(list(row))
    wb.close()
    return headers, rows, date_cols


def carregar_routescape(filepath):
    """
    Retorna (wb, ws, headers, rows, date_cols).
    - headers/rows/date_cols vêm do cache (leitura rápida)
    - wb/ws são carregados fresh SÓ quando é preciso escrever

    Para leitura pura, prefira chamar _carregar_routescape_dados()
    diretamente (não instancia openpyxl em modo editável).
    """
    mtime = os.path.getmtime(filepath) if os.path.exists(filepath) else 0.0
    headers, rows, date_cols = _carregar_routescape_dados(filepath, mtime)
    wb = load_workbook(filepath)
    ws = wb[ROUTESCAPE_SHEET]
    return wb, ws, headers, rows, date_cols


def invalidar_cache_routescape():
    """Chamada após salvar a planilha — força releitura na próxima chamada."""
    _carregar_routescape_dados.clear()


def encontrar_coluna_mes(date_cols, mes, ano):
    for col_idx, dt in date_cols:
        if dt.month == mes and dt.year == ano:
            return col_idx
    return None


def adicionar_coluna_mes(wb, ws, date_cols, mes, ano):
    if not date_cols:
        return None
    last_col_idx, _ = date_cols[-1]
    nova_col_excel = last_col_idx + 2
    ws.cell(row=1, column=nova_col_excel, value=datetime(ano, mes, 26))
    return last_col_idx + 1


def ler_dados_mes_importado(filepath, mes, ano):
    """
    Lê a coluna do mês/ano na planilha histórica e retorna os
    varejistas que já têm valor importado (units > 0).
    Retorna (target_col, [ {account, address, city, units}, ... ])
    ou (None, []) se a coluna não existe ainda.

    Usa a versão cacheada (_carregar_routescape_dados) — não abre
    o workbook em modo editável.
    """
    mtime = os.path.getmtime(filepath) if os.path.exists(filepath) else 0.0
    headers, rows, date_cols = _carregar_routescape_dados(filepath, mtime)
    target_col = encontrar_coluna_mes(date_cols, mes, ano)
    if target_col is None:
        return None, []

    importados = []
    for row in rows:
        if target_col >= len(row):
            continue
        val = row[target_col]
        if val is None or (isinstance(val, (int, float)) and val == 0):
            continue
        acc  = str(row[COL_ACCOUNT]).strip()  if row[COL_ACCOUNT] else ""
        addr = str(row[COL_ADDRESS]).strip()  if row[COL_ADDRESS] else ""
        city = str(row[COL_CITY]).strip()     if row[COL_CITY]    else ""
        if not acc or acc.upper() == "ACCOUNT":
            continue
        importados.append({
            "Varejista": acc,
            "Endereço" : addr,
            "Cidade"   : city,
            "Units"    : int(val) if isinstance(val, (int, float)) else val,
        })
    return target_col, importados


def excluir_dados_mes(filepath, mes, ano):
    """
    Remove (zera) todos os valores da coluna do mês/ano
    na planilha histórica. Retorna quantas linhas foram limpas.
    """
    wb, ws, headers, rows, date_cols = carregar_routescape(filepath)
    target_col = encontrar_coluna_mes(date_cols, mes, ano)
    if target_col is None:
        return 0

    limpos = 0
    # +2 porque: linha 1 do Excel é cabeçalho, e openpyxl indexa a partir de 1
    for i, row in enumerate(rows):
        if target_col < len(row) and row[target_col] is not None:
            ws.cell(row=i + 2, column=target_col + 1, value=None)
            limpos += 1
    wb.save(filepath)
    invalidar_cache_routescape()
    return limpos


def ler_arquivo_mensal(filepath):
    """
    Lê arquivo mensal da distribuidora. Tenta identificar:
      - Período: procura "thru" em qualquer lugar das 10 primeiras linhas,
                  ou padrão de data MM/DD/YYYY isolado
      - Varejistas: account (col A), address (col B), city (col C),
                    item (col D), units (col E).
                    Cada varejista tem 1 linha agregadora (col D == "Total")
                    e várias linhas de produtos individuais. Lemos apenas
                    a linha "Total" de cada varejista para evitar dupla
                    contagem.
    """
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active
    mes, ano = None, None

    # ── 1) Procura o período: "thru" nas 15 primeiras linhas ─────────────
    for linha_num in range(1, 16):
        for cell in ws[linha_num]:
            if not cell.value:
                continue
            texto = str(cell.value)
            if "thru" in texto.lower():
                m = re.search(r"(\d+)/\d+/(\d{4})", texto)
                if m:
                    mes, ano = int(m.group(1)), int(m.group(2))
                    break
        if mes and ano:
            break

    # ── 2) Fallback: procura qualquer data MM/DD/YYYY nas 15 primeiras linhas ─
    if not (mes and ano):
        for linha_num in range(1, 16):
            for cell in ws[linha_num]:
                if not cell.value:
                    continue
                texto = str(cell.value)
                m = re.search(r"\b(\d{1,2})/\d{1,2}/(\d{4})\b", texto)
                if m:
                    mes_cand = int(m.group(1))
                    ano_cand = int(m.group(2))
                    if 1 <= mes_cand <= 12 and 2020 <= ano_cand <= 2035:
                        mes, ano = mes_cand, ano_cand
                        break
            if mes and ano:
                break

    # ── 3) Lê varejistas: auto-detecta linha inicial ──────────────────────
    # Formato esperado:
    #   col A = Retail Account, col B = Address, col C = City,
    #   col D = Item Name (ou "Total" para a linha agregadora do varejista),
    #   col E = Units Sold
    # Cada varejista tem 1 linha "Total" + N linhas de produtos. Lemos só
    # as linhas "Total" para pegar o agregado direto (evita dupla contagem).
    contas = []
    for row in ws.iter_rows(min_row=1, values_only=True):
        if len(row) < 5:
            continue
        acc, addr, _col_c, item, units = row[0], row[1], row[2], row[3], row[4]

        # Pula linhas de cabeçalho/grand total
        acc_txt = str(acc).strip() if acc else ""
        if not acc_txt or acc_txt.upper() in ("", "TOTAL", "ACCOUNT", "GRAND TOTAL", "RETAIL ACCOUNTS"):
            continue

        # Pega só a linha agregadora do varejista (col D == "Total")
        item_txt = str(item).strip().lower() if item is not None else ""
        if item_txt != "total":
            continue

        # Precisa ter units numérico
        if not isinstance(units, (int, float)):
            continue

        contas.append({
            "account": acc_txt.upper(),
            "address": str(addr).strip().upper() if addr else "",
            "units"  : int(units),
        })

    return mes, ano, contas


# ─────────────────────────────────────────
#  FUNÇÃO — EXPORTAR PLANOS DE AÇÃO PARA EXCEL
# ─────────────────────────────────────────

def exportar_planos_excel(com_acao: list, mes: int, ano: int) -> bytes:
    """
    Gera um arquivo Excel com todos os planos de ação do mês.
    Cada campo do formulário de visita vira uma coluna.
    Retorna bytes para download direto no Streamlit.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Planos {MESES_PT[mes]} {ano}"

    # ── Estilo do cabeçalho ──────────────────────────────
    header_fill   = PatternFill("solid", fgColor="1B4332")   # verde escuro Babe
    header_font   = Font(bold=True, color="FFFFFF", size=11)
    center_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin_border   = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    fill_pending  = PatternFill("solid", fgColor="FFF3CD")   # amarelo suave
    fill_done     = PatternFill("solid", fgColor="D1E7DD")   # verde suave

    # ── Colunas ──────────────────────────────────────────
    colunas = [
        ("Varejista",            30),
        ("Endereço",             35),
        ("Data da Visita",       16),
        ("Visita Realizada",     16),
        ("Notas da Visita",      40),
        ("Plano de Ação",        50),
        ("Follow-up",            40),
        ("Encaminhado",          14),
        ("Data Encaminhamento",  20),
    ]

    # Cabeçalho
    for col_idx, (titulo, largura) in enumerate(colunas, start=1):
        cell = ws.cell(row=1, column=col_idx, value=titulo)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center_align
        cell.border    = thin_border
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(col_idx)
        ].width = largura

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    # ── Dados ─────────────────────────────────────────────
    for row_idx, item in enumerate(com_acao, start=2):
        enc_str  = "Sim ✅" if item.get("forwarded") else "Não ⏳"
        vis_str  = "Sim ✅" if item.get("visited")   else "Não ⬜"
        enc_data = (item.get("forwarded_date") or "")[:10]

        valores = [
            item.get("account", ""),
            item.get("address", ""),
            item.get("date", "") or "",
            vis_str,
            item.get("notes", "") or "",
            item.get("action", "") or "",
            item.get("followup", "") or "",
            enc_str,
            enc_data,
        ]

        row_fill = fill_done if item.get("forwarded") else fill_pending

        for col_idx, valor in enumerate(valores, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=valor)
            cell.fill      = row_fill
            cell.border    = thin_border
            cell.alignment = left_align if col_idx in (1, 2, 5, 6, 7) else center_align

        ws.row_dimensions[row_idx].height = 45

    # Filtro automático
    ws.auto_filter.ref = ws.dimensions

    # Retorna bytes
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────
#  COMPONENTE — FORMULÁRIO DE VISITA
#  Reutilizado na página Top 55
# ─────────────────────────────────────────

def formulario_visita(acc: str, addr: str, city: str, mes: int, ano: int, key_prefix: str):
    """Renderiza o formulário de registro de visita e salva ao clicar."""
    v = get_visita(mes, ano, acc, addr)

    c1, c2 = st.columns(2)
    with c1:
        visited = st.checkbox(
            "✅ Visita realizada",
            value=bool(v.get("visited", False)),
            key=f"vis_{key_prefix}",
        )
        try:
            data_default = datetime.strptime(v["date"], "%Y-%m-%d").date() if v.get("date") else None
        except Exception:
            data_default = None

        visit_date = st.date_input(
            "📅 Data da visita",
            value=data_default,
            key=f"dat_{key_prefix}",
        )
        has_action = st.radio(
            "📌 Há plano de ação?",
            ["Não", "Sim"],
            index=1 if v.get("has_action") else 0,
            horizontal=True,
            key=f"hac_{key_prefix}",
        )

    with c2:
        notes = st.text_area(
            "📝 Notas da visita",
            value=v.get("notes", ""),
            height=80,
            key=f"not_{key_prefix}",
        )
        action = st.text_area(
            "🎯 Plano de ação (detalhes)",
            value=v.get("action", ""),
            height=80,
            key=f"act_{key_prefix}",
        )
        followup = st.text_area(
            "🔁 Follow-up / observações",
            value=v.get("followup", ""),
            height=60,
            key=f"flw_{key_prefix}",
        )

    if st.button("💾 Salvar visita", key=f"sav_{key_prefix}", type="primary"):
        salvar_visita(mes, ano, acc, addr, {
            "visited"       : visited,
            "date"          : visit_date.isoformat() if visit_date else None,
            "notes"         : notes,
            "has_action"    : (has_action == "Sim"),
            "action"        : action,
            "followup"      : followup,
            "forwarded"     : v.get("forwarded", False),      # preserva status de encaminhamento
            "forwarded_date": v.get("forwarded_date", None),
        })
        st.success("✅ Visita salva!")
        st.rerun()


# ─────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────
st.sidebar.title("🍵 Babe Kombucha")
st.sidebar.markdown("#### Gestão de Visitas ao Varejo")
st.sidebar.divider()

hoje = datetime.now()
mes_sel = st.sidebar.selectbox(
    "📅 Mês de referência:",
    list(MESES_PT.keys()),
    index=hoje.month - 1,
    format_func=lambda m: MESES_PT[m],
)
ano_sel = st.sidebar.number_input(
    "Ano:", min_value=2024, max_value=2030, value=hoje.year, step=1
)
st.sidebar.divider()

pagina = st.sidebar.radio(
    "Navegar para:",
    [
        "📥  Importar Dados Mensais",
        "🏆  Top 55 + Registro de Visitas",
        "📤  Planos de Ação",
    ],
    label_visibility="collapsed",
)

# ── Painel de persistência (sidebar) ───────────────────────
st.sidebar.divider()
st.sidebar.markdown("**💾 Dados salvos**")

# 1) Meses importados na planilha histórica (usa cache — sem abrir xlsx editável)
_hist_opts_sb = encontrar_arquivo_historico()
if _hist_opts_sb:
    try:
        _hp = os.path.join(APP_DIR, _hist_opts_sb[0])
        _mt = os.path.getmtime(_hp) if os.path.exists(_hp) else 0.0
        _, _, _dc = _carregar_routescape_dados(_hp, _mt)
        if _dc:
            meses_tags = " · ".join(d.strftime("%b/%y") for _, d in _dc[-8:])
            st.sidebar.caption(f"📊 Planilha: **{meses_tags}**")
        else:
            st.sidebar.caption("📊 Planilha: nenhum mês importado ainda")
    except Exception:
        st.sidebar.caption("📊 Planilha: erro ao ler")
else:
    st.sidebar.caption("📊 Planilha: arquivo não encontrado")

# 2) Visitas registradas no JSON
_vis_data = load_visitas()
if _vis_data:
    _total_vis = sum(len(v) for v in _vis_data.values())
    _meses_vis = sorted(_vis_data.keys())
    _vis_tags  = " · ".join(_meses_vis[-4:])
    st.sidebar.caption(f"📝 Visitas: **{_total_vis}** registros ({_vis_tags})")
else:
    st.sidebar.caption("📝 Visitas: nenhuma registrada ainda")

st.sidebar.caption(
    f"📂 Pasta: `...retail_platform/`\n\n"
    "✅ **visitas.json** — registros de visitas (permanente)\n\n"
    "✅ **planilha .xlsx** — dados de vendas (permanente)\n\n"
    "⚠️ O ranking Top 55 precisa ser recalculado uma vez por sessão — "
    "**mas os dados já estão salvos**."
)


# ═══════════════════════════════════════════════════════════
#  PÁGINA 1 — IMPORTAR DADOS MENSAIS
# ═══════════════════════════════════════════════════════════
if "📥" in pagina:
    st.title(f"📥 Importar Dados Mensais — {MESES_PT[mes_sel]} {ano_sel}")

    hist_options = encontrar_arquivo_historico()
    if not hist_options:
        st.error("⚠️ Nenhuma planilha .xlsx encontrada nesta pasta.")
        st.stop()

    arquivo_historico = st.selectbox("📁 Planilha histórica (John Lenore):", hist_options)
    hist_path = os.path.join(APP_DIR, arquivo_historico)

    # ── Verificar se o mês selecionado já foi importado ───────────────────
    target_col_existente, importados = ler_dados_mes_importado(hist_path, mes_sel, ano_sel)

    if importados:
        # ══════ MÊS JÁ TEM DADOS IMPORTADOS ═══════════════════════════════
        st.success(
            f"✅ **{MESES_PT[mes_sel]} {ano_sel}** já está importado — "
            f"**{len(importados)} varejistas** com valores registrados."
        )

        df_imp = pd.DataFrame(importados)
        total_units = df_imp["Units"].sum() if "Units" in df_imp.columns else 0
        c1, c2 = st.columns(2)
        c1.metric("Varejistas importados", len(df_imp))
        c2.metric("Total de unidades",     f"{int(total_units):,}")

        with st.expander("🔍 Ver todos os dados importados deste mês", expanded=False):
            st.dataframe(df_imp, use_container_width=True)

        st.divider()
        st.markdown("### 🔁 Precisa reimportar ou corrigir?")
        st.caption(
            "Se a importação deste mês foi indevida ou veio com valores errados, "
            "use **Excluir dados** para zerar a coluna e depois refaça o upload."
        )

        col_del1, col_del2 = st.columns([1, 2])
        confirmar_del = col_del2.checkbox(
            f"Confirmo que quero excluir os dados de {MESES_PT[mes_sel]} {ano_sel}",
            key=f"conf_del_{mes_sel}_{ano_sel}",
        )
        if col_del1.button("🗑️ Excluir dados deste mês", type="secondary", disabled=not confirmar_del):
            limpos = excluir_dados_mes(hist_path, mes_sel, ano_sel)
            st.success(f"✅ {limpos} linhas zeradas. O mês **{MESES_PT[mes_sel]} {ano_sel}** foi limpo da planilha.")
            st.rerun()

        st.divider()
        st.markdown("### ⬆️ Sobrescrever com novo upload (opcional)")

    else:
        # ══════ MÊS AINDA NÃO IMPORTADO ═══════════════════════════════════
        st.info(
            f"Nenhum dado importado ainda para **{MESES_PT[mes_sel]} {ano_sel}**. "
            "Faça o upload do arquivo mensal abaixo."
        )

    # ── Uploader (sempre visível) ─────────────────────────────────────────
    upload_mensal = st.file_uploader(
        "📂 Arquivo mensal (BABE DIST BY ACCTS...):",
        type=["xlsx"],
        key=f"up_{mes_sel}_{ano_sel}",
    )

    if upload_mensal:
        temp_path = os.path.join(APP_DIR, "_temp_mensal.xlsx")
        with open(temp_path, "wb") as f:
            f.write(upload_mensal.read())

        # ── Tenta ler o arquivo — se quebrar, mostra o erro ──────────────
        try:
            mes, ano, contas_mensais = ler_arquivo_mensal(temp_path)
        except Exception as _err:
            st.error(f"❌ Erro ao ler o arquivo: {_err}")
            mes, ano, contas_mensais = None, None, []

        # ── DIAGNÓSTICO SEMPRE VISÍVEL ───────────────────────────────────
        st.markdown("#### 📋 Diagnóstico do arquivo")
        diag1, diag2, diag3 = st.columns(3)
        diag1.metric("Mês detectado", MESES_PT[mes] if mes else "—")
        diag2.metric("Ano detectado", ano if ano else "—")
        diag3.metric("Varejistas lidos", len(contas_mensais))

        # ── Se não detectou automaticamente, permite escolher manualmente ─
        if not mes or not ano:
            st.warning(
                "⚠️ Não consegui identificar o período automaticamente "
                "(o leitor procura a palavra **'thru'** na linha 4 do arquivo). "
                "Selecione o período manualmente abaixo:"
            )
            col_m1, col_m2 = st.columns(2)
            mes = col_m1.selectbox(
                "Mês deste arquivo:",
                list(MESES_PT.keys()),
                index=mes_sel - 1,
                format_func=lambda m: MESES_PT[m],
                key="manual_mes",
            )
            ano = col_m2.number_input(
                "Ano:", min_value=2024, max_value=2030,
                value=ano_sel, step=1, key="manual_ano",
            )

            # Se também não leu varejistas, tenta mostrar preview cru
            if not contas_mensais:
                st.info(
                    "Também não encontrei linhas de varejistas no formato esperado "
                    "(account na coluna A, address na B, units na D, começando na linha 8). "
                    "Abaixo está o conteúdo cru das primeiras linhas para conferência:"
                )
                try:
                    raw = pd.read_excel(temp_path, header=None, nrows=15)
                    st.dataframe(raw, use_container_width=True)
                except Exception as _er2:
                    st.error(f"Não consegui nem fazer o preview cru: {_er2}")

        # ── Se bateu tudo, segue para importação ──────────────────────────
        if mes and ano and contas_mensais:
            if (mes, ano) != (mes_sel, ano_sel):
                st.warning(
                    f"⚠️ O arquivo é de **{MESES_PT[mes]} {ano}**, mas você selecionou "
                    f"**{MESES_PT[mes_sel]} {ano_sel}** na barra lateral. "
                    "Ajuste a seleção ou confira o arquivo."
                )

            st.info(
                f"📅 Período: **{MESES_PT[mes]} {ano}** — "
                f"**{len(contas_mensais)} varejistas** prontos para importar."
            )

            with st.expander("🔍 Pré-visualizar primeiras 10 linhas lidas"):
                st.dataframe(pd.DataFrame(contas_mensais[:10]))

            wb, ws, headers, rows, date_cols = carregar_routescape(hist_path)
            target_col = encontrar_coluna_mes(date_cols, mes, ano)

            if target_col is None:
                st.warning(
                    f"⚠️ A coluna de **{MESES_PT[mes]}/{ano}** ainda não existe "
                    "na planilha histórica."
                )
                if st.button("➕ Criar nova coluna para este mês"):
                    target_col = adicionar_coluna_mes(wb, ws, date_cols, mes, ano)
                    wb.save(hist_path)
                    invalidar_cache_routescape()
                    st.success(
                        f"✅ Coluna criada para {MESES_PT[mes]}/{ano}! "
                        "Clique em IMPORTAR abaixo."
                    )
                    st.rerun()
            else:
                st.success("✅ Coluna de destino localizada na planilha histórica.")

                if st.button("🚀 IMPORTAR DADOS AGORA", type="primary"):
                    hist_index = {}
                    for i, row in enumerate(rows):
                        acc  = str(row[COL_ACCOUNT]).strip().upper() if row[COL_ACCOUNT] else ""
                        addr = str(row[COL_ADDRESS]).strip().upper() if row[COL_ADDRESS] else ""
                        hist_index[(acc, addr)] = i

                    atualizados, nao_encontrados = 0, []
                    for item in contas_mensais:
                        chave = (item["account"], item["address"])
                        if chave in hist_index:
                            linha_excel = hist_index[chave] + 2
                            ws.cell(row=linha_excel, column=target_col + 1, value=item["units"])
                            atualizados += 1
                        else:
                            nao_encontrados.append(item)

                    wb.save(hist_path)
                    invalidar_cache_routescape()
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

                    st.success(f"✅ **{atualizados} varejistas atualizados** com sucesso!")
                    if nao_encontrados:
                        st.warning(f"⚠️ {len(nao_encontrados)} contas não encontradas no histórico.")
                        with st.expander("Ver contas não encontradas"):
                            st.dataframe(pd.DataFrame(nao_encontrados))


# ═══════════════════════════════════════════════════════════
#  PÁGINA 2 — TOP 55 + REGISTRO DE VISITAS
# ═══════════════════════════════════════════════════════════
elif "🏆" in pagina:
    st.title(f"🏆 Top 55 Varejistas — {MESES_PT[mes_sel]} {ano_sel}")
    st.markdown(
        "Ranking pelos últimos meses de vendas. "
        "Clique em qualquer varejista para registrar a visita do mês selecionado."
    )

    hist_options = encontrar_arquivo_historico()
    if not hist_options:
        st.error("⚠️ Nenhuma planilha encontrada.")
        st.stop()

    arquivo_historico = st.selectbox("📁 Planilha histórica:", hist_options)

    col_sl1, col_sl2 = st.columns(2)
    n_meses    = col_sl1.slider("Meses para o ranking:", 6, 24, 18)
    max_por_dia = col_sl2.slider("Máx. visitas por dia:", 5, 20, 15)

    # Chave única por mês/ano — evita crossover entre meses
    _sk     = f"top55_{mes_sel}_{ano_sel}"
    _sk_n   = f"top55_n_{mes_sel}_{ano_sel}"
    _sk_dia = f"top55_dia_{mes_sel}_{ano_sel}"

    if st.button("📊 Calcular Top 55", type="primary"):
        with st.spinner("Calculando..."):
            _hist_path = os.path.join(APP_DIR, arquivo_historico)
            _mt_h      = os.path.getmtime(_hist_path) if os.path.exists(_hist_path) else 0.0
            _, rows, date_cols = _carregar_routescape_dados(_hist_path, _mt_h)
            ultimos = date_cols[-n_meses:] if len(date_cols) >= n_meses else date_cols

            varejistas = []
            for row in rows:
                acc = row[COL_ACCOUNT]
                if not acc or str(acc).strip() in ("", "ACCOUNT"):
                    continue
                total = sum(
                    row[ci] for ci, _ in ultimos
                    if ci < len(row) and isinstance(row[ci], (int, float))
                )
                if total > 0:
                    varejistas.append({
                        "Varejista": str(acc).strip(),
                        "Endereço" : str(row[COL_ADDRESS]).strip() if row[COL_ADDRESS] else "",
                        "Cidade"   : str(row[COL_CITY]).strip()    if row[COL_CITY]    else "",
                        "ZIP"      : str(row[COL_ZIPCODE]).strip()  if row[COL_ZIPCODE] else "",
                        "Zona"     : str(row[COL_ZONE]).strip()     if row[COL_ZONE]    else "",
                        f"Total {n_meses}m": total,
                    })

            df = (
                pd.DataFrame(varejistas)
                .sort_values(f"Total {n_meses}m", ascending=False)
                .head(55)
                .reset_index(drop=True)
            )
            df.index += 1
            st.session_state[_sk]     = df
            st.session_state[_sk_n]   = n_meses
            st.session_state[_sk_dia] = max_por_dia

    # ── Exibe resultado ──────────────────────────────────────────────────────
    if _sk not in st.session_state:
        st.info(
            f"Clique em **📊 Calcular Top 55** para gerar o ranking de "
            f"**{MESES_PT[mes_sel]} {ano_sel}**. "
            "Cada mês precisa ser calculado uma vez por sessão."
        )
    if _sk in st.session_state:
        df   = st.session_state[_sk]
        n    = st.session_state[_sk_n]
        mpd  = st.session_state.get(_sk_dia, max_por_dia)

        # Métricas
        c1, c2, c3 = st.columns(3)
        c1.metric("Varejistas no ranking", len(df))
        c2.metric(f"Total vendido (Top 55)", f"{df[f'Total {n}m'].sum():,}")
        c3.metric("Média por varejista", f"{df[f'Total {n}m'].mean():.0f}")

        # Agrupamento geográfico real (ZIP + jitter + K-Means)
        with st.spinner("📍 Calculando agrupamento geográfico..."):
            df_geo = agrupar_por_proximidade(df, mpd)
        n_dias = int(df_geo["Dia"].max())

        st.divider()
        st.markdown(
            f"**{n_dias} dias de visitas** — "
            f"mês de referência: **{MESES_PT[mes_sel]} {ano_sel}**"
        )

        # Carregar visitas do mês
        visitas_mes = load_visitas().get(mes_key(mes_sel, ano_sel), {})

        # Uma aba por dia + aba de lista completa
        tab_labels = [f"📆 Dia {d}" for d in range(1, n_dias + 1)] + ["📋 Lista Completa"]
        tabs = st.tabs(tab_labels)

        for d in range(1, n_dias + 1):
            with tabs[d - 1]:
                dia_df = df_geo[df_geo["Dia"] == d].reset_index(drop=True)
                st.markdown(f"**{len(dia_df)} visitas neste dia**")

                for _, row in dia_df.iterrows():
                    acc  = row["Varejista"]
                    addr = row["Endereço"]
                    city = row["Cidade"]
                    vk   = visita_key(acc, addr)
                    v    = visitas_mes.get(vk, {})

                    # Ícones de status
                    s_icon = "✅" if v.get("visited")    else "⬜"
                    a_icon = "📌" if v.get("has_action") and not v.get("forwarded") else (
                             "✉️" if v.get("forwarded")  else ""
                    )
                    total  = row[f"Total {n}m"]

                    label = f"{s_icon}{a_icon} **{acc}** — {addr}, {city} | {total:,} un."

                    with st.expander(label, expanded=False):
                        kp = f"d{d}_{acc[:15]}_{addr[:10]}_{mes_sel}{ano_sel}"
                        formulario_visita(acc, addr, city, mes_sel, ano_sel, kp)

        # Aba lista completa
        with tabs[-1]:
            st.dataframe(df, use_container_width=True)

        # ── Exportar planos de ação do mês ──────────────────────────────────
        st.divider()
        _vis_mes_export = load_visitas().get(mes_key(mes_sel, ano_sel), {})
        _com_acao_export = []
        for _vk, _v in _vis_mes_export.items():
            if _v.get("has_action") and _v.get("action", "").strip():
                _partes = _vk.split("|", 1)
                _com_acao_export.append({
                    "account"       : _partes[0] if _partes else _vk,
                    "address"       : _partes[1] if len(_partes) > 1 else "",
                    "action"        : _v.get("action", ""),
                    "notes"         : _v.get("notes", ""),
                    "followup"      : _v.get("followup", ""),
                    "date"          : _v.get("date", ""),
                    "visited"       : _v.get("visited", False),
                    "forwarded"     : _v.get("forwarded", False),
                    "forwarded_date": _v.get("forwarded_date", ""),
                })

        col_xls, col_xls_info = st.columns([1, 3])
        with col_xls:
            if _com_acao_export:
                _xls_bytes  = exportar_planos_excel(_com_acao_export, mes_sel, ano_sel)
                _xls_nome   = f"Planos_Acao_{MESES_PT[mes_sel]}_{ano_sel}.xlsx"
                st.download_button(
                    label=f"📥 Exportar Planos de Ação ({len(_com_acao_export)})",
                    data=_xls_bytes,
                    file_name=_xls_nome,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            else:
                st.info("Nenhum plano de ação salvo neste mês ainda.")
        with col_xls_info:
            if _com_acao_export:
                st.caption(
                    f"Exporta **{len(_com_acao_export)} planos de ação** de "
                    f"**{MESES_PT[mes_sel]} {ano_sel}** em Excel — "
                    "cada campo do formulário vira uma coluna."
                )


# ═══════════════════════════════════════════════════════════
#  PÁGINA 3 — PLANOS DE AÇÃO
# ═══════════════════════════════════════════════════════════
elif "📤" in pagina:
    st.title(f"📤 Planos de Ação — {MESES_PT[mes_sel]} {ano_sel}")
    st.markdown(
        "Varejistas com plano de ação registrado no mês selecionado. "
        "Marque como **encaminhado** após enviar para o time comercial / produção."
    )

    visitas      = load_visitas()
    mk           = mes_key(mes_sel, ano_sel)
    mes_visitas  = visitas.get(mk, {})

    # Separar os que têm plano de ação
    com_acao = []
    for vk, v in mes_visitas.items():
        if v.get("has_action") and v.get("action", "").strip():
            partes = vk.split("|", 1)
            com_acao.append({
                "account"       : partes[0] if partes else vk,
                "address"       : partes[1] if len(partes) > 1 else "",
                "action"        : v.get("action", ""),
                "notes"         : v.get("notes", ""),
                "followup"      : v.get("followup", ""),
                "date"          : v.get("date", ""),
                "visited"       : v.get("visited", False),
                "forwarded"     : v.get("forwarded", False),
                "forwarded_date": v.get("forwarded_date", ""),
                "vk"            : vk,
            })

    if not com_acao:
        st.info(
            f"Nenhum plano de ação registrado para **{MESES_PT[mes_sel]} {ano_sel}**. "
            "Registre as visitas na página 🏆 Top 55."
        )
        st.stop()

    pendentes     = [x for x in com_acao if not x["forwarded"]]
    encaminhados  = [x for x in com_acao if x["forwarded"]]

    # ── EXPORTAR EXCEL — TODOS DE UMA VEZ ─────────────────────────────────
    st.subheader("📤 Exportar para Excel")
    col_e1, col_e2 = st.columns([1, 1])
    with col_e1:
        try:
            _excel_bytes_all = exportar_planos_excel(com_acao, mes_sel, ano_sel)
            _excel_nome_all  = f"Planos_Acao_{MESES_PT[mes_sel]}_{ano_sel}_TODOS.xlsx"
            st.download_button(
                label=f"⬇️ Baixar TODOS os planos ({len(com_acao)})",
                data=_excel_bytes_all,
                file_name=_excel_nome_all,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        except Exception as _e:
            st.error(f"Erro ao gerar Excel: {_e}")

    # ── EXPORTAÇÃO EM LOTE (SELECIONAR QUAIS) ─────────────────────────────
    with col_e2:
        with st.expander(f"☑️ Selecionar e exportar em lote ({len(com_acao)} disponíveis)"):
            st.caption("Marque os planos que deseja exportar num único arquivo Excel.")

            # Botões "marcar todos / nenhum"
            cb1, cb2 = st.columns(2)
            if cb1.button("✅ Marcar todos", key="sel_all_planos", use_container_width=True):
                for item in com_acao:
                    st.session_state[f"sel_{item['vk']}"] = True
                st.rerun()
            if cb2.button("⬜ Desmarcar todos", key="sel_none_planos", use_container_width=True):
                for item in com_acao:
                    st.session_state[f"sel_{item['vk']}"] = False
                st.rerun()

            selecionados = []
            for item in com_acao:
                status_ico = "✉️" if item["forwarded"] else "📌"
                checked = st.checkbox(
                    f"{status_ico} {item['account']} — {item['address'][:40]}",
                    key=f"sel_{item['vk']}",
                    value=st.session_state.get(f"sel_{item['vk']}", False),
                )
                if checked:
                    selecionados.append(item)

            if selecionados:
                _excel_bytes_sel = exportar_planos_excel(selecionados, mes_sel, ano_sel)
                _excel_nome_sel  = (
                    f"Planos_Acao_{MESES_PT[mes_sel]}_{ano_sel}_"
                    f"SELECIONADOS_{len(selecionados)}.xlsx"
                )
                st.download_button(
                    label=f"⬇️ Baixar {len(selecionados)} selecionados",
                    data=_excel_bytes_sel,
                    file_name=_excel_nome_sel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.caption("_Nenhum item selecionado ainda._")

    st.divider()

    # Métricas
    m1, m2, m3 = st.columns(3)
    m1.metric("Total com plano de ação", len(com_acao))
    m2.metric("⏳ Pendentes",   len(pendentes))
    m3.metric("✉️ Encaminhados", len(encaminhados))

    # ── Pendentes ───────────────────────────────────────────────────────────
    if pendentes:
        st.divider()
        st.subheader("⏳ Pendentes de Encaminhamento")

        if st.button(
            f"✉️ Marcar TODOS como encaminhados ({len(pendentes)})",
            type="primary",
        ):
            visitas = load_visitas()
            agora   = datetime.now().isoformat()
            for item in pendentes:
                if mk in visitas and item["vk"] in visitas[mk]:
                    visitas[mk][item["vk"]]["forwarded"]      = True
                    visitas[mk][item["vk"]]["forwarded_date"] = agora
            save_visitas(visitas)
            st.success("✅ Todos marcados como encaminhados!")
            st.rerun()

        for item in pendentes:
            with st.expander(f"📌 **{item['account']}** — {item['address']}"):
                col_i, col_a = st.columns([1, 2])
                col_i.markdown(f"**Data da visita:** {item['date'] or '—'}")
                col_i.markdown(f"**Visitado:** {'Sim ✅' if item['visited'] else 'Não ⬜'}")
                col_a.markdown(f"**Plano de ação:**\n\n{item['action']}")
                if item["notes"]:
                    col_a.markdown(f"**Notas:** {item['notes']}")

                btn1, btn2 = st.columns(2)

                with btn1:
                    if st.button(
                        "✉️ Marcar como encaminhado",
                        key=f"enc_{item['vk']}",
                        use_container_width=True,
                    ):
                        visitas = load_visitas()
                        if mk in visitas and item["vk"] in visitas[mk]:
                            visitas[mk][item["vk"]]["forwarded"]      = True
                            visitas[mk][item["vk"]]["forwarded_date"] = datetime.now().isoformat()
                        save_visitas(visitas)
                        st.success("✅ Encaminhado!")
                        st.rerun()

                # ── EXPORTAÇÃO INDIVIDUAL DESTE PLANO ─────────────────────
                with btn2:
                    _bytes_ind = exportar_planos_excel([item], mes_sel, ano_sel)
                    _nome_ind  = (
                        f"Plano_{item['account'][:30].replace(' ', '_')}_"
                        f"{MESES_PT[mes_sel]}_{ano_sel}.xlsx"
                    )
                    st.download_button(
                        label="⬇️ Exportar só este plano",
                        data=_bytes_ind,
                        file_name=_nome_ind,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_ind_{item['vk']}",
                        use_container_width=True,
                    )

    # ── Já encaminhados ─────────────────────────────────────────────────────
    if encaminhados:
        st.divider()
        st.subheader("✉️ Já Encaminhados")
        for item in encaminhados:
            data_enc = item["forwarded_date"][:10] if item["forwarded_date"] else "?"
            with st.expander(f"✉️ {item['account']} — encaminhado em {data_enc}"):
                st.markdown(f"**Plano de ação:** {item['action']}")
                if item["notes"]:
                    st.markdown(f"**Notas:** {item['notes']}")

                bt1, bt2 = st.columns(2)

                with bt1:
                    # Permite desfazer o encaminhamento se necessário
                    if st.button(
                        "↩️ Desfazer encaminhamento",
                        key=f"unenc_{item['vk']}",
                        use_container_width=True,
                    ):
                        visitas = load_visitas()
                        if mk in visitas and item["vk"] in visitas[mk]:
                            visitas[mk][item["vk"]]["forwarded"]      = False
                            visitas[mk][item["vk"]]["forwarded_date"] = None
                        save_visitas(visitas)
                        st.rerun()

                with bt2:
                    _bytes_ind_e = exportar_planos_excel([item], mes_sel, ano_sel)
                    _nome_ind_e  = (
                        f"Plano_{item['account'][:30].replace(' ', '_')}_"
                        f"{MESES_PT[mes_sel]}_{ano_sel}.xlsx"
                    )
                    st.download_button(
                        label="⬇️ Exportar só este plano",
                        data=_bytes_ind_e,
                        file_name=_nome_ind_e,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_ind_enc_{item['vk']}",
                        use_container_width=True,
                    )
