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
import re
from datetime import datetime, date
import os
import json

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

# ─────────────────────────────────────────
#  FUNÇÕES — VISITAS (JSON)
#  Cada mês é uma "fatia" independente.
#  Meses passados ficam congelados.
# ─────────────────────────────────────────

def load_visitas() -> dict:
    """Carrega o arquivo visitas.json (cria vazio se não existir)."""
    if os.path.exists(VISITAS_FILE):
        with open(VISITAS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_visitas(data: dict):
    """Salva o dicionário de visitas no arquivo JSON."""
    with open(VISITAS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)


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
    arquivos = [f for f in os.listdir(APP_DIR) if f.endswith(".xlsx")]
    lenore = [f for f in arquivos if "lenore" in f.lower() or "john" in f.lower()]
    return lenore if lenore else arquivos


def carregar_routescape(filepath):
    wb = load_workbook(filepath)
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
    return wb, ws, headers, rows, date_cols


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


def ler_arquivo_mensal(filepath):
    wb = load_workbook(filepath)
    ws = wb.active
    mes, ano = None, None
    for cell in ws[4]:
        if cell.value and "thru" in str(cell.value):
            m = re.match(r"(\d+)/\d+/(\d{4})", str(cell.value).strip())
            if m:
                mes, ano = int(m.group(1)), int(m.group(2))
            break
    contas = []
    for row in ws.iter_rows(min_row=8, values_only=True):
        acc, addr, units = row[0], row[1], row[3]
        if acc and str(acc).strip() not in ("", "Total") and units is not None:
            contas.append({
                "account": str(acc).strip().upper(),
                "address": str(addr).strip().upper() if addr else "",
                "units"  : int(units) if isinstance(units, (int, float)) else 0,
            })
    return mes, ano, contas


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


# ═══════════════════════════════════════════════════════════
#  PÁGINA 1 — IMPORTAR DADOS MENSAIS
# ═══════════════════════════════════════════════════════════
if "📥" in pagina:
    st.title("📥 Importar Dados Mensais")
    st.markdown(
        "Faça o upload do arquivo mensal de distribuição. "
        "A plataforma detecta o mês automaticamente e atualiza a planilha histórica."
    )

    hist_options = encontrar_arquivo_historico()
    if not hist_options:
        st.error("⚠️ Nenhuma planilha .xlsx encontrada nesta pasta.")
        st.stop()

    arquivo_historico = st.selectbox("📁 Planilha histórica (John Lenore):", hist_options)
    upload_mensal     = st.file_uploader("📂 Arquivo mensal (BABE DIST BY ACCTS...):", type=["xlsx"])

    if upload_mensal:
        temp_path = os.path.join(APP_DIR, "_temp_mensal.xlsx")
        with open(temp_path, "wb") as f:
            f.write(upload_mensal.read())

        mes, ano, contas_mensais = ler_arquivo_mensal(temp_path)

        if not mes or not ano:
            st.error("❌ Não foi possível identificar o período no arquivo.")
            st.stop()

        st.info(f"📅 Período detectado: **{MESES_PT[mes]} {ano}** — **{len(contas_mensais)} varejistas** encontrados.")

        with st.expander("🔍 Pré-visualizar arquivo mensal (primeiras 10 linhas)"):
            st.dataframe(pd.DataFrame(contas_mensais[:10]))

        wb, ws, headers, rows, date_cols = carregar_routescape(os.path.join(APP_DIR, arquivo_historico))
        target_col = encontrar_coluna_mes(date_cols, mes, ano)

        if target_col is None:
            st.warning(f"⚠️ A coluna de **{MESES_PT[mes]}/{ano}** ainda não existe na planilha histórica.")
            if st.button("➕ Criar nova coluna para este mês"):
                target_col = adicionar_coluna_mes(wb, ws, date_cols, mes, ano)
                wb.save(os.path.join(APP_DIR, arquivo_historico))
                st.success(f"✅ Coluna criada para {MESES_PT[mes]}/{ano}! Clique em IMPORTAR abaixo.")
                st.rerun()
        else:
            st.success("✅ Coluna de destino localizada na planilha histórica.")

        if target_col is not None:
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

                wb.save(os.path.join(APP_DIR, arquivo_historico))
                os.remove(temp_path)

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

    if st.button("📊 Calcular Top 55", type="primary"):
        with st.spinner("Calculando..."):
            _, _, _, rows, date_cols = carregar_routescape(os.path.join(APP_DIR, arquivo_historico))
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
            st.session_state["top55_df"] = df
            st.session_state["top55_n"]  = n_meses
            st.session_state["top55_dia"] = max_por_dia

    # ── Exibe resultado ──────────────────────────────────────────────────────
    if "top55_df" in st.session_state:
        df   = st.session_state["top55_df"]
        n    = st.session_state["top55_n"]
        mpd  = st.session_state.get("top55_dia", max_por_dia)

        # Métricas
        c1, c2, c3 = st.columns(3)
        c1.metric("Varejistas no ranking", len(df))
        c2.metric(f"Total vendido (Top 55)", f"{df[f'Total {n}m'].sum():,}")
        c3.metric("Média por varejista", f"{df[f'Total {n}m'].mean():.0f}")

        # Ordenação geográfica → dias
        df_geo = df.sort_values(["Zona", "ZIP"]).reset_index(drop=True)
        df_geo["Dia"] = (df_geo.index // mpd) + 1
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

                if st.button(
                    "✉️ Marcar como encaminhado",
                    key=f"enc_{item['vk']}",
                ):
                    visitas = load_visitas()
                    if mk in visitas and item["vk"] in visitas[mk]:
                        visitas[mk][item["vk"]]["forwarded"]      = True
                        visitas[mk][item["vk"]]["forwarded_date"] = datetime.now().isoformat()
                    save_visitas(visitas)
                    st.success("✅ Encaminhado!")
                    st.rerun()

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

                # Permite desfazer o encaminhamento se necessário
                if st.button("↩️ Desfazer encaminhamento", key=f"unenc_{item['vk']}"):
                    visitas = load_visitas()
                    if mk in visitas and item["vk"] in visitas[mk]:
                        visitas[mk][item["vk"]]["forwarded"]      = False
                        visitas[mk][item["vk"]]["forwarded_date"] = None
                    save_visitas(visitas)
                    st.rerun()
