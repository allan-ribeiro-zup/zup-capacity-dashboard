import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import io

st.set_page_config(
    page_title="Capacidade · Zup",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background: #F7F4F2; }
  [data-testid="stSidebar"] { background: #2C1A1A; }
  [data-testid="stSidebar"] * { color: #F7F4F2 !important; }
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label { color: #F7F4F2 !important; }
  .metric-card {
      background: white; border-radius: 12px; padding: 18px 22px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08); text-align: center;
      margin-bottom: 4px;
  }
  .metric-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 1px; }
  .metric-value { font-size: 30px; font-weight: 700; color: #2C1A1A; margin-top: 4px; }
  .metric-value.orange { color: #E05A2B; }
  .metric-value.green  { color: #27AE60; }
  .metric-value.red    { color: #E74C3C; }
  .metric-value.blue   { color: #2980B9; }
  .section-title {
      font-size: 17px; font-weight: 700; color: #2C1A1A;
      border-left: 4px solid #E05A2B; padding-left: 12px;
      margin: 20px 0 10px;
  }
  .stDataFrame { border-radius: 8px; overflow: hidden; }

  /* ── Estrutura das Squads ── */
  .squad-wrap { font-family: sans-serif; }
  .squad-dir {
      display: inline-block; background: #E05A2B; color: #fff;
      border-radius: 8px; padding: 6px 18px; font-size: 13px;
      font-weight: 600; text-align: center; margin: 4px;
  }
  .squad-dir small { display: block; font-size: 10px; opacity: .8; font-weight: 400; }
  .squad-coord {
      display: inline-block; background: #2C1A1A; color: #fff;
      border-radius: 6px; padding: 4px 14px; font-size: 12px;
      font-weight: 500; text-align: center; margin: 3px;
  }
  .squad-coord small { display: block; font-size: 10px; opacity: .7; }
  .squad-cap-block {
      border: 1px solid #e0dcd8; border-radius: 10px;
      padding: 10px 12px; background: #FAF8F6;
      margin-bottom: 8px;
  }
  .squad-cap-title {
      font-size: 12px; font-weight: 700; color: #2C1A1A;
      border-bottom: 2px solid #E05A2B; padding-bottom: 4px;
      margin-bottom: 8px;
  }
  .sq-pill {
      display: inline-block; color: #fff; font-size: 10px;
      font-weight: 600; padding: 2px 10px; border-radius: 4px;
      margin-bottom: 4px;
  }
  .membro-tag {
      display: inline-block; font-size: 10px; padding: 1px 7px;
      border: 1px solid #ddd; border-radius: 4px;
      background: #fff; color: #333; margin: 1px;
  }
  .membro-tag b { color: #888; font-weight: 400; }
  .membro-tag.hl { background: #FEF9C3; border-color: #CA8A04; color: #713F12; }
</style>
""", unsafe_allow_html=True)

SPRINTS = ["Sprint Planning","Sprint 1","Sprint 2","Sprint 3","Sprint 4","Sprint 5","Sprint 6"]
STATUS_COLORS = {
    "Concluído":    "#27AE60",
    "Em Andamento": "#F1C40F",
    "Não Iniciado": "#95A5A6",
    "Bloqueado":    "#E74C3C",
    "Cancelado":    "#BDC3C7",
    "Em Atraso":    "#E05A2B",
    "Despriorizado":"#D5D8DC",
}

GITHUB_URL        = "https://raw.githubusercontent.com/allan-ribeiro-zup/zup-capacity-dashboard/main/template_capacity_zup.xlsx"
GITHUB_DEPLOY_URL = "https://raw.githubusercontent.com/allan-ribeiro-zup/zup-capacity-dashboard/main/template_deploy_zup.xlsx"

# ── Sidebar ───────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚡ Capacidade Zup")
    st.markdown("---")
    if st.button("🔄 Atualizar dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")
    pagina = st.radio("Navegar", [
        "🏠 Dashboard",
        "🗺️ Roadmap Detalhado",
        "📊 Atividades por Release",
        "👥 Planejamento das Squads",
        "📅 Férias e Ausências",
        "👤 Gestão de Pessoas",
        "🚀 Acompanhamento de Deploy",
        "🏢 Estrutura das Squads",
    ])
    st.markdown("---")
    st.caption("Zup · Gestão de Capacidade · 2026")

# ── Helpers ───────────────────────────────────────────────────────
def card(label, value, color=""):
    cls = f"metric-value {color}".strip()
    return f"""<div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="{cls}">{value}</div>
    </div>"""

def safe_pct(df, col, val):
    if col not in df.columns or df.empty:
        return 0
    return len(df[df[col] == val])

# ── Carregamento: capacidade ──────────────────────────────────────
@st.cache_data(show_spinner="Carregando dados do GitHub...", ttl=300)
def load():
    try:
        resp = requests.get(GITHUB_URL, timeout=30)
        resp.raise_for_status()
        file_content = io.BytesIO(resp.content)
        xls   = pd.ExcelFile(file_content, engine="openpyxl")
        names = xls.sheet_names
        def get(name):
            if name in names:
                df = xls.parse(name, header=1)
                df.columns = [str(c).strip() for c in df.columns]
                unnamed = sum(1 for c in df.columns if "Unnamed" in str(c))
                if unnamed > len(df.columns) // 2:
                    df = xls.parse(name, header=0)
                    df.columns = [str(c).strip() for c in df.columns]
                df = df.dropna(how="all")
                return df
            return pd.DataFrame()
        def get_ferias():
            nm = "Férias e Ausências" if "Férias e Ausências" in names else "Ausências"
            if nm not in names:
                return pd.DataFrame()
            df = xls.parse(nm, header=2)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.dropna(how="all")
            if "Zupper" in df.columns:
                df = df[df["Zupper"].notna() & (df["Zupper"] != "Zupper")]
            return df
        return {
            "membros":    get("Membros"),
            "ausencias":  get_ferias(),
            "cap":        get("Capacidade por Sprint"),
            "roadmap":    get("Roadmap"),
            "atividades": get("Atividades por Release"),
            "ok":         True,
        }
    except Exception as e:
        return {"ok": False, "erro": str(e)}

# ── Carregamento: deploy ──────────────────────────────────────────
@st.cache_data(show_spinner="Carregando dados de deploy...", ttl=300)
def load_deploy():
    try:
        resp = requests.get(GITHUB_DEPLOY_URL, timeout=30)
        resp.raise_for_status()
        xls  = pd.ExcelFile(io.BytesIO(resp.content), engine="openpyxl")
        abas = [s for s in xls.sheet_names if s.isdigit() and len(s) == 8]
        abas_sorted = sorted(abas, reverse=True)
        sheets = {}
        for aba in abas_sorted:
            df = xls.parse(aba, header=2)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.dropna(how="all")
            df = df[df["BUG"].notna() & (df["BUG"].astype(str).str.strip() != "")]
            sheets[aba] = df
        return {"ok": True, "sheets": sheets, "abas": abas_sorted}
    except Exception as e:
        return {"ok": False, "erro": str(e), "sheets": {}, "abas": []}

# ── Carrega dados ─────────────────────────────────────────────────
with st.spinner("Carregando dados do GitHub..."):
    data = load()

if not data.get("ok"):
    st.error(f"❌ Erro ao carregar planilha: {data.get('erro','')}")
    st.info("Verifique se o arquivo 'template_capacity_zup.xlsx' está no repositório GitHub.")
    st.stop()

roadmap   = data["roadmap"]
ativ      = data["atividades"]
cap       = data["cap"]
membros   = data["membros"]
ausencias = data["ausencias"]


# ════════════════════════════════════════════════════════════════
# 🏠 DASHBOARD
# ════════════════════════════════════════════════════════════════
if pagina == "🏠 Dashboard":
    st.markdown("# Dashboard de Capacidade · Gestão de Portfólio")

    squads_all   = sorted(roadmap["Squad"].dropna().unique().tolist()) if "Squad" in roadmap.columns else []
    releases_all = sorted(ativ["Release"].dropna().unique().tolist())  if "Release" in ativ.columns else []

    cf1, cf2 = st.columns([3,1])
    with cf1:
        squad_sel = st.multiselect("Filtrar Squad", squads_all, default=squads_all)
    with cf2:
        rel_sel = st.selectbox("Release", ["Todas"] + releases_all)

    rm = roadmap.copy()
    if squad_sel and "Squad" in rm.columns:
        rm = rm[rm["Squad"].isin(squad_sel)]

    total   = len(rm)
    andando = safe_pct(rm, "Status", "Em Andamento")
    concl   = safe_pct(rm, "Status", "Concluído")
    bloq    = safe_pct(rm, "Status", "Bloqueado")
    canc    = safe_pct(rm, "Status", "Cancelado")
    pct_med = rm["% Conclusão"].mean() if "% Conclusão" in rm.columns and not rm.empty else 0

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.markdown(card("Total Features",   total,           "orange"), unsafe_allow_html=True)
    c2.markdown(card("Em Andamento",     andando                  ), unsafe_allow_html=True)
    c3.markdown(card("Concluídas",       concl,           "green" ), unsafe_allow_html=True)
    c4.markdown(card("Bloqueadas",       bloq,            "red"   ), unsafe_allow_html=True)
    c5.markdown(card("Canceladas",       canc                     ), unsafe_allow_html=True)
    c6.markdown(card("% Conclusão Méd.", f"{pct_med:.0f}%","blue" ), unsafe_allow_html=True)

    st.markdown("---")
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown('<div class="section-title">Distribuição por Status</div>', unsafe_allow_html=True)
        if "Status" in rm.columns and not rm.empty:
            sc = rm["Status"].value_counts().reset_index()
            sc.columns = ["Status","Count"]
            fig = px.pie(sc, values="Count", names="Status", hole=0.52,
                         color="Status", color_discrete_map=STATUS_COLORS)
            fig.update_traces(textposition="outside", textinfo="label+value")
            fig.update_layout(showlegend=False, margin=dict(t=20,b=20,l=20,r=120), height=320)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem dados de status no roadmap.")

    with col_b:
        st.markdown('<div class="section-title">Horas Disponíveis por Squad</div>', unsafe_allow_html=True)
        if not cap.empty and "Squad" in cap.columns:
            sp_cols = [c for c in SPRINTS if c in cap.columns]
            if sp_cols:
                cap_sq = cap.groupby("Squad")[sp_cols].sum()
                cap_sq["Total"] = cap_sq.sum(axis=1)
                cap_sq = cap_sq.reset_index().sort_values("Total", ascending=True)
                fig2 = px.bar(cap_sq, x="Total", y="Squad", orientation="h",
                              color_discrete_sequence=["#E05A2B"],
                              labels={"Total":"Horas","Squad":""})
                fig2.update_layout(height=320, plot_bgcolor="white",
                                   paper_bgcolor="rgba(0,0,0,0)", margin=dict(t=10,b=10,l=10,r=10))
                st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Sem dados de capacidade.")

    st.markdown('<div class="section-title">Capacidade Total por Sprint</div>', unsafe_allow_html=True)
    if not cap.empty:
        sp_cols = [c for c in SPRINTS if c in cap.columns]
        if sp_cols:
            totais = cap[sp_cols].sum()
            fig3 = go.Figure()
            fig3.add_trace(go.Bar(x=sp_cols, y=totais.values,
                                  marker_color="#E05A2B", name="Capacidade (h)", opacity=0.85))
            fig3.update_layout(height=280, plot_bgcolor="white",
                                paper_bgcolor="rgba(0,0,0,0)",
                                margin=dict(t=10,b=10,l=0,r=0),
                                yaxis_title="Horas")
            st.plotly_chart(fig3, use_container_width=True)

    st.markdown('<div class="section-title">Utilização por Plataforma × Sprint</div>', unsafe_allow_html=True)
    if not cap.empty and "Plataforma" in cap.columns:
        sp_cols = [c for c in SPRINTS if c in cap.columns]
        if sp_cols:
            pivot = cap.groupby("Plataforma")[sp_cols].sum()
            fig4 = px.imshow(pivot, text_auto=".0f",
                             color_continuous_scale=["#FFEEEE","#F1C40F","#27AE60"],
                             labels=dict(x="Sprint", y="Plataforma", color="Horas"),
                             aspect="auto")
            fig4.update_layout(height=220, margin=dict(t=10,b=10), coloraxis_showscale=False)
            st.plotly_chart(fig4, use_container_width=True)


# ════════════════════════════════════════════════════════════════
# 🗺️ ROADMAP DETALHADO
# ════════════════════════════════════════════════════════════════
elif pagina == "🗺️ Roadmap Detalhado":
    st.markdown("# Roadmap Detalhado")
    st.caption("Visão completa com todas as features, tarefas, responsáveis e cronograma.")

    if roadmap.empty:
        st.warning("Sem dados na aba 'Roadmap' do Excel."); st.stop()

    f1,f2,f3,f4 = st.columns(4)
    squads_r = sorted(roadmap["Squad"].dropna().unique())     if "Squad"      in roadmap.columns else []
    status_r = sorted(roadmap["Status"].dropna().unique())    if "Status"     in roadmap.columns else []
    plats_r  = sorted(roadmap["Plataforma"].dropna().unique()) if "Plataforma" in roadmap.columns else []

    with f1: sq  = st.multiselect("Squad",      squads_r)
    with f2: st_ = st.multiselect("Status",     status_r)
    with f3: pl  = st.multiselect("Plataforma", plats_r)
    with f4: bsc = st.text_input("🔍 Buscar tarefa")

    rm2 = roadmap.copy()
    if sq  and "Squad"      in rm2.columns: rm2 = rm2[rm2["Squad"].isin(sq)]
    if st_ and "Status"     in rm2.columns: rm2 = rm2[rm2["Status"].isin(st_)]
    if pl  and "Plataforma" in rm2.columns: rm2 = rm2[rm2["Plataforma"].isin(pl)]
    if bsc and "Tarefa"     in rm2.columns: rm2 = rm2[rm2["Tarefa"].str.contains(bsc, case=False, na=False)]

    k1,k2,k3,k4,k5 = st.columns(5)
    k1.markdown(card("Total", len(rm2), "orange"), unsafe_allow_html=True)
    k2.markdown(card("Em Andamento", safe_pct(rm2,"Status","Em Andamento")), unsafe_allow_html=True)
    k3.markdown(card("Concluídas",   safe_pct(rm2,"Status","Concluído"), "green"), unsafe_allow_html=True)
    k4.markdown(card("Bloqueadas",   safe_pct(rm2,"Status","Bloqueado"), "red"),   unsafe_allow_html=True)
    pct2 = rm2["% Conclusão"].mean() if "% Conclusão" in rm2.columns and not rm2.empty else 0
    k5.markdown(card("% Médio", f"{pct2:.0f}%", "blue"), unsafe_allow_html=True)

    st.markdown(f"**{len(rm2)} itens** | use os filtros acima para refinar")

    bg_map = {"Concluído":"#C8E6C9","Em Andamento":"#FFF9C4",
              "Bloqueado":"#FFCDD2","Cancelado":"#EEEEEE","Não Iniciado":"#F5F5F5"}

    def color_st(val):
        return f"background-color:{bg_map.get(val,'')}"

    show_cols = [c for c in ["Nº Feature","Tipo","Tema","Entrega de Valor","Tarefa",
                              "Squad","Plataforma","Prioridade","Status",
                              "Sprint Início","Sprint Fim","% Conclusão","Horas Estimadas"]
                 if c in rm2.columns]

    styled = rm2[show_cols].style.map(color_st, subset=["Status"] if "Status" in show_cols else [])
    st.dataframe(styled, use_container_width=True, height=500)

    st.markdown("---")
    ca, cb = st.columns(2)
    with ca:
        st.markdown('<div class="section-title">Features por Squad</div>', unsafe_allow_html=True)
        if "Squad" in rm2.columns and not rm2.empty:
            sq_cnt = rm2["Squad"].value_counts().reset_index()
            sq_cnt.columns = ["Squad","Count"]
            fig_sq = px.bar(sq_cnt, x="Count", y="Squad", orientation="h",
                            color_discrete_sequence=["#2C1A1A"])
            fig_sq.update_layout(height=250, plot_bgcolor="white",
                                 paper_bgcolor="rgba(0,0,0,0)", margin=dict(t=5,b=5))
            st.plotly_chart(fig_sq, use_container_width=True)
    with cb:
        st.markdown('<div class="section-title">Features por Plataforma</div>', unsafe_allow_html=True)
        if "Plataforma" in rm2.columns and not rm2.empty:
            pl_cnt = rm2["Plataforma"].value_counts().reset_index()
            pl_cnt.columns = ["Plataforma","Count"]
            fig_pl = px.pie(pl_cnt, values="Count", names="Plataforma", hole=0.4,
                            color_discrete_sequence=["#E05A2B","#2C1A1A","#F1C40F","#27AE60"])
            fig_pl.update_layout(height=250, margin=dict(t=5,b=5,l=5,r=80))
            st.plotly_chart(fig_pl, use_container_width=True)


# ════════════════════════════════════════════════════════════════
# 📊 ATIVIDADES POR RELEASE
# ════════════════════════════════════════════════════════════════
elif pagina == "📊 Atividades por Release":
    st.markdown("# Roadmap de Atividades por Release")

    if ativ.empty:
        st.warning("Sem dados na aba 'Atividades por Release'."); st.stop()

    releases = sorted(ativ["Release"].dropna().unique()) if "Release" in ativ.columns else []
    if not releases:
        st.warning("Coluna 'Release' não encontrada."); st.stop()

    rel = st.selectbox("Selecionar Release", releases)
    df_rel = ativ[ativ["Release"] == rel].copy()

    squads_rel = sorted(df_rel["Squad"].dropna().unique()) if "Squad" in df_rel.columns else []
    sq_fil = st.multiselect("Filtrar Squad", squads_rel, default=squads_rel)
    if sq_fil and "Squad" in df_rel.columns:
        df_rel = df_rel[df_rel["Squad"].isin(sq_fil)]

    n = len(df_rel)
    st.caption(f"{n} atividades · {df_rel['Grande Iniciativa'].nunique() if 'Grande Iniciativa' in df_rel.columns else '?'} iniciativas · ordenado por prioridade")

    sp_cols = [c for c in SPRINTS if c in df_rel.columns]

    st.markdown('<div class="section-title">Progresso % por Iniciativa × Sprint</div>', unsafe_allow_html=True)
    if sp_cols and "Grande Iniciativa" in df_rel.columns and not df_rel.empty:
        pivot = df_rel.set_index("Grande Iniciativa")[sp_cols].apply(pd.to_numeric, errors="coerce")
        fig_h = px.imshow(
            pivot, text_auto=".0f", zmin=0, zmax=100,
            color_continuous_scale=["#EEEEEE","#F1C40F","#27AE60"],
            aspect="auto",
            labels=dict(x="Sprint", y="Iniciativa", color="%")
        )
        fig_h.update_layout(height=max(220, n*38), margin=dict(t=10,b=10,l=0,r=0),
                            coloraxis_showscale=False)
        st.plotly_chart(fig_h, use_container_width=True)

    st.markdown('<div class="section-title">Status por Iniciativa</div>', unsafe_allow_html=True)
    if "Grande Iniciativa" in df_rel.columns and "Status" in df_rel.columns:
        for _, row in df_rel.iterrows():
            ini    = row.get("Grande Iniciativa","—")
            status = row.get("Status","Não Iniciado")
            squad  = row.get("Squad","")
            vals = [row.get(s) for s in sp_cols if pd.notna(row.get(s)) and isinstance(row.get(s),(int,float))]
            pct  = max(vals) if vals else 0
            col_i, col_p = st.columns([4,6])
            with col_i:
                st.markdown(f"**{ini}** <span style='font-size:11px;color:#888'>· {squad}</span>", unsafe_allow_html=True)
                st.caption(status)
            with col_p:
                st.progress(int(pct)/100, text=f"{int(pct)}%")

    with st.expander("📋 Ver tabela completa"):
        st.dataframe(df_rel, use_container_width=True, hide_index=True)


# ════════════════════════════════════════════════════════════════
# 👥 PLANEJAMENTO DAS SQUADS
# ════════════════════════════════════════════════════════════════
elif pagina == "👥 Planejamento das Squads":
    st.markdown("# Planejamento das Squads")

    if cap.empty:
        st.warning("Sem dados na aba 'Capacidade por Sprint'."); st.stop()

    sp_cols = [c for c in SPRINTS if c in cap.columns]

    squads_c = sorted(cap["Squad"].dropna().unique()) if "Squad" in cap.columns else []
    sq_sel   = st.multiselect("Filtrar Squad", squads_c, default=squads_c)
    cap_f    = cap[cap["Squad"].isin(sq_sel)] if sq_sel and "Squad" in cap.columns else cap

    plats = sorted(cap_f["Plataforma"].dropna().unique()) if "Plataforma" in cap_f.columns else []
    if plats:
        cols_k = st.columns(len(plats))
        for idx, plat in enumerate(plats):
            df_p  = cap_f[cap_f["Plataforma"]==plat]
            total = df_p[sp_cols].sum().sum() if sp_cols else 0
            n_mem = len(membros[membros["Plataforma"]==plat]) if not membros.empty and "Plataforma" in membros.columns else "—"
            with cols_k[idx]:
                st.markdown(f"""<div class="metric-card">
                    <div class="metric-label">{plat}</div>
                    <div class="metric-value orange">{total:.0f}h</div>
                    <div style="font-size:11px;color:#888">{n_mem} membros</div>
                </div>""", unsafe_allow_html=True)

    st.markdown("---")

    st.markdown('<div class="section-title">Capacidade por Sprint e Plataforma</div>', unsafe_allow_html=True)
    if sp_cols and "Plataforma" in cap_f.columns:
        grp  = cap_f.groupby("Plataforma")[sp_cols].sum().reset_index()
        melt = grp.melt(id_vars="Plataforma", var_name="Sprint", value_name="Horas")
        fig_b = px.bar(melt, x="Sprint", y="Horas", color="Plataforma", barmode="group",
                       color_discrete_sequence=["#E05A2B","#2C1A1A","#F1C40F","#27AE60"],
                       category_orders={"Sprint": sp_cols})
        fig_b.update_layout(height=340, plot_bgcolor="white",
                            paper_bgcolor="rgba(0,0,0,0)", margin=dict(t=10,b=10))
        st.plotly_chart(fig_b, use_container_width=True)

    st.markdown('<div class="section-title">Heatmap de Horas por Plataforma × Sprint</div>', unsafe_allow_html=True)
    if sp_cols and "Plataforma" in cap_f.columns:
        pvt = cap_f.groupby("Plataforma")[sp_cols].sum()
        fig_hm = px.imshow(pvt, text_auto=".0f",
                           color_continuous_scale=["#FFEEEE","#F1C40F","#27AE60"],
                           aspect="auto",
                           labels=dict(x="Sprint", y="Plataforma", color="Horas"))
        fig_hm.update_layout(height=230, margin=dict(t=5,b=5), coloraxis_showscale=False)
        st.plotly_chart(fig_hm, use_container_width=True)

    st.markdown('<div class="section-title">Membros por Squad</div>', unsafe_allow_html=True)
    if not membros.empty and "Squad" in membros.columns:
        for sq in (sq_sel if sq_sel else squads_c):
            df_sq = membros[membros["Squad"]==sq]
            if not df_sq.empty:
                with st.expander(f"👥 {sq} — {len(df_sq)} membros"):
                    show = [c for c in ["Nome","Plataforma","Senioridade","Horas/Sprint",
                                        "% Run (Sustentação)","% Treinamento"] if c in df_sq.columns]
                    st.dataframe(df_sq[show], use_container_width=True, hide_index=True)
    else:
        st.info("Preencha a aba 'Membros' no Excel para ver os detalhes do time.")


# ════════════════════════════════════════════════════════════════
# 📅 FÉRIAS E AUSÊNCIAS
# ════════════════════════════════════════════════════════════════
elif pagina == "📅 Férias e Ausências":
    st.markdown("# Férias e Ausências 2026")

    aus = ausencias.copy()
    if aus.empty:
        st.warning("Sem dados na aba 'Férias e Ausências'."); st.stop()

    MESES_NOMES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                   "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]

    cols      = list(aus.columns)
    total_col = cols[-1]

    projetos_disp = sorted(aus["Projeto"].dropna().replace("",float("nan")).dropna().unique()) if "Projeto" in aus.columns else []
    perfis_disp   = sorted(aus["Perfil"].dropna().unique()) if "Perfil" in aus.columns else []

    cf1, cf2, cf3 = st.columns(3)
    with cf1: proj_f  = st.multiselect("Projeto", projetos_disp)
    with cf2: perf_f  = st.multiselect("Perfil",  perfis_disp)
    with cf3:
        mes_idx = st.selectbox("Mês de referência", range(12),
                               format_func=lambda x: MESES_NOMES[x])

    aus_f = aus.copy()
    if proj_f and "Projeto" in aus_f.columns:
        aus_f = aus_f[aus_f["Projeto"].isin(proj_f)]
    if perf_f and "Perfil" in aus_f.columns:
        aus_f = aus_f[aus_f["Perfil"].isin(perf_f)]

    total_zuppers = len(aus_f)
    sem_projeto   = len(aus_f[aus_f["Projeto"].fillna("").str.strip() == ""]) if "Projeto" in aus_f.columns else 0

    col_dias_mes = None
    offset = 6 + mes_idx * 3 + 2
    if offset < len(cols):
        col_dias_mes = cols[offset]

    ausentes_mes = 0
    if col_dias_mes:
        ausentes_mes = len(aus_f[pd.to_numeric(aus_f[col_dias_mes], errors="coerce").fillna(0) > 0])

    total_dias_agendados = pd.to_numeric(aus_f[total_col], errors="coerce").sum() if total_col in aus_f.columns else 0

    k1,k2,k3,k4 = st.columns(4)
    k1.markdown(card("Total Zuppers",       total_zuppers),             unsafe_allow_html=True)
    k2.markdown(card("Ausentes em " + MESES_NOMES[mes_idx], ausentes_mes, "orange"), unsafe_allow_html=True)
    k3.markdown(card("Disponíveis p/ Realocação", sem_projeto, "red"),  unsafe_allow_html=True)
    k4.markdown(card("Total Dias Agendados", f"{int(total_dias_agendados)}d", "blue"), unsafe_allow_html=True)

    st.markdown("---")

    st.markdown('<div class="section-title">Ausências por Mês</div>', unsafe_allow_html=True)
    dados_meses = []
    for i, mes in enumerate(MESES_NOMES):
        offset_d = 6 + i * 3 + 2
        if offset_d < len(cols):
            col_d = cols[offset_d]
            n_ausentes = len(aus_f[pd.to_numeric(aus_f[col_d], errors="coerce").fillna(0) > 0])
            total_d    = pd.to_numeric(aus_f[col_d], errors="coerce").fillna(0).sum()
            dados_meses.append({"Mês": mes, "Pessoas Ausentes": n_ausentes, "Total Dias": int(total_d)})

    df_meses = pd.DataFrame(dados_meses)
    if not df_meses.empty:
        fig_m = go.Figure()
        fig_m.add_trace(go.Bar(
            x=df_meses["Mês"], y=df_meses["Pessoas Ausentes"],
            name="Pessoas Ausentes", marker_color="#E05A2B", opacity=0.85
        ))
        fig_m.add_trace(go.Scatter(
            x=df_meses["Mês"], y=df_meses["Total Dias"],
            name="Total de Dias", yaxis="y2",
            line=dict(color="#2C1A1A", width=2), mode="lines+markers"
        ))
        fig_m.update_layout(
            height=320, plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=20,b=20),
            yaxis=dict(title="Pessoas Ausentes"),
            yaxis2=dict(title="Total de Dias", overlaying="y", side="right"),
            legend=dict(orientation="h", y=1.1)
        )
        st.plotly_chart(fig_m, use_container_width=True)

    st.markdown('<div class="section-title">Disponibilidade por Zupper × Mês</div>', unsafe_allow_html=True)
    if "Zupper" in aus_f.columns:
        heat_data = []
        for _, row in aus_f.iterrows():
            zupper = row.get("Zupper", "")
            for i, mes in enumerate(MESES_NOMES):
                offset_d = 6 + i * 3 + 2
                if offset_d < len(cols):
                    val = pd.to_numeric(row.get(cols[offset_d], 0), errors="coerce")
                    heat_data.append({"Zupper": zupper, "Mês": mes, "Dias": 0 if pd.isna(val) else int(val)})

        df_heat = pd.DataFrame(heat_data)
        if not df_heat.empty:
            pivot = df_heat.pivot_table(index="Zupper", columns="Mês", values="Dias", fill_value=0)
            pivot = pivot[MESES_NOMES] if all(m in pivot.columns for m in MESES_NOMES) else pivot
            fig_h = px.imshow(
                pivot, text_auto=True, zmin=0, zmax=30,
                color_continuous_scale=["#FFFFFF","#BDD7EE","#E05A2B"],
                aspect="auto",
                labels=dict(x="Mês", y="Zupper", color="Dias")
            )
            fig_h.update_layout(
                height=max(300, len(aus_f)*28),
                margin=dict(t=10,b=10), coloraxis_showscale=True
            )
            st.plotly_chart(fig_h, use_container_width=True)

    st.markdown('<div class="section-title">Calendário Anual de Férias (Gantt)</div>', unsafe_allow_html=True)
    if "Zupper" in aus_f.columns:
        import datetime
        gantt_data = []
        for _, row in aus_f.iterrows():
            zupper  = str(row.get("Zupper",""))
            projeto = str(row.get("Projeto","Sem projeto"))
            for i, mes in enumerate(MESES_NOMES):
                off_ini  = 6 + i * 3
                off_fim  = off_ini + 1
                off_dias = off_ini + 2
                if off_dias < len(cols):
                    ini_v  = row.get(cols[off_ini],  "")
                    fim_v  = row.get(cols[off_fim],  "")
                    dias_v = pd.to_numeric(row.get(cols[off_dias], 0), errors="coerce")
                    if pd.notna(dias_v) and dias_v > 0 and ini_v and str(ini_v).strip():
                        try:
                            ano     = 2026
                            ini_str = f"{str(ini_v).strip()}/{ano}"
                            fim_str = f"{str(fim_v).strip()}/{ano}" if fim_v and str(fim_v).strip() else ini_str
                            ini_dt  = datetime.datetime.strptime(ini_str, "%d/%m/%Y")
                            fim_dt  = datetime.datetime.strptime(fim_str, "%d/%m/%Y")
                            gantt_data.append({
                                "Zupper":  zupper,
                                "Projeto": projeto,
                                "Início":  ini_dt,
                                "Fim":     fim_dt,
                                "Dias":    int(dias_v),
                                "Mês":     mes,
                            })
                        except:
                            pass

        if gantt_data:
            df_gantt = pd.DataFrame(gantt_data)
            fig_g = px.timeline(
                df_gantt, x_start="Início", x_end="Fim", y="Zupper",
                color="Projeto", hover_data=["Dias","Mês"],
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig_g.update_yaxes(autorange="reversed")
            fig_g.update_layout(
                height=max(300, len(aus_f)*32),
                plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                margin=dict(t=10,b=10),
                xaxis=dict(range=["2026-01-01","2026-12-31"], dtick="M1", tickformat="%b")
            )
            st.plotly_chart(fig_g, use_container_width=True)
        else:
            st.info("Preencha as datas de início e fim na planilha para visualizar o Gantt.")

    with st.expander("📋 Ver tabela completa"):
        show_cols = [c for c in ["Zupper","Projeto","Perfil","Day Off 1","Day Off 2","Aniversário", total_col] if c in aus_f.columns]
        st.dataframe(aus_f[show_cols], use_container_width=True, hide_index=True)


# ════════════════════════════════════════════════════════════════
# 👤 GESTÃO DE PESSOAS
# ════════════════════════════════════════════════════════════════
elif pagina == "👤 Gestão de Pessoas":
    st.markdown("# Gestão de Pessoas")

    aus = ausencias.copy()
    if aus.empty:
        st.warning("Sem dados na aba 'Férias e Ausências'."); st.stop()

    MESES_NOMES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                   "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    cols      = list(aus.columns)
    total_col = cols[-1]

    st.markdown("---")

    st.markdown('<div class="section-title">⚠️ Zuppers Disponíveis para Realocação</div>', unsafe_allow_html=True)
    if "Projeto" in aus.columns:
        aus["Projeto"] = aus["Projeto"].fillna("")
        sem_proj = aus[aus["Projeto"].astype(str).str.strip() == ""].copy()
        if not sem_proj.empty:
            show = [c for c in ["Zupper","Perfil","Day Off 1","Day Off 2", total_col] if c in sem_proj.columns]
            st.dataframe(
                sem_proj[show].style.map(lambda _: "background-color:#FFD9B3", subset=["Zupper"] if "Zupper" in show else []),
                use_container_width=True, hide_index=True
            )
        else:
            st.success("Todos os Zuppers estão alocados em projetos.")

    st.markdown("---")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<div class="section-title">Headcount por Projeto</div>', unsafe_allow_html=True)
        if "Projeto" in aus.columns:
            aus["Projeto"] = aus["Projeto"].fillna("")
            hc     = aus[aus["Projeto"].astype(str).str.strip() != ""]
            hc_cnt = hc["Projeto"].value_counts().reset_index()
            hc_cnt.columns = ["Projeto","Pessoas"]
            fig_hc = px.bar(hc_cnt, x="Pessoas", y="Projeto", orientation="h",
                            color_discrete_sequence=["#2C1A1A"], text="Pessoas")
            fig_hc.update_traces(textposition="outside")
            fig_hc.update_layout(height=320, plot_bgcolor="white",
                                 paper_bgcolor="rgba(0,0,0,0)",
                                 margin=dict(t=10,b=10,l=10,r=40))
            st.plotly_chart(fig_hc, use_container_width=True)

    with col_b:
        st.markdown('<div class="section-title">Distribuição por Perfil Técnico</div>', unsafe_allow_html=True)
        if "Perfil" in aus.columns:
            perf_cnt = aus["Perfil"].value_counts().reset_index()
            perf_cnt.columns = ["Perfil","Pessoas"]
            fig_p = px.pie(perf_cnt, values="Pessoas", names="Perfil", hole=0.45,
                           color_discrete_sequence=px.colors.qualitative.Set2)
            fig_p.update_layout(height=320, margin=dict(t=10,b=10,l=10,r=80))
            st.plotly_chart(fig_p, use_container_width=True)

    st.markdown('<div class="section-title">Perfis por Projeto</div>', unsafe_allow_html=True)
    if "Projeto" in aus.columns and "Perfil" in aus.columns:
        proj_perf = aus[aus["Projeto"].astype(str).str.strip() != ""].groupby(
            ["Projeto","Perfil"]).size().reset_index(name="Pessoas")
        fig_pp = px.bar(proj_perf, x="Projeto", y="Pessoas", color="Perfil",
                        barmode="stack", text="Pessoas",
                        color_discrete_sequence=px.colors.qualitative.Set2)
        fig_pp.update_traces(textposition="inside")
        fig_pp.update_layout(height=340, plot_bgcolor="white",
                             paper_bgcolor="rgba(0,0,0,0)", margin=dict(t=10,b=10))
        st.plotly_chart(fig_pp, use_container_width=True)

    st.markdown('<div class="section-title">Concentração de Ausências por Mês</div>', unsafe_allow_html=True)
    dados_conc = []
    for i, mes in enumerate(MESES_NOMES):
        offset_d = 6 + i * 3 + 2
        if offset_d < len(cols):
            col_d   = cols[offset_d]
            por_proj = aus[aus["Projeto"].fillna("").str.strip() != ""].copy()
            por_proj["dias_num"] = pd.to_numeric(por_proj[col_d], errors="coerce").fillna(0)
            n_aus   = len(por_proj[por_proj["dias_num"] > 0])
            total_p = len(por_proj)
            pct     = round(n_aus / total_p * 100, 1) if total_p > 0 else 0
            dados_conc.append({"Mês": mes, "% Ausentes": pct, "Pessoas": n_aus})

    df_conc = pd.DataFrame(dados_conc)
    if not df_conc.empty:
        fig_c = px.bar(df_conc, x="Mês", y="% Ausentes",
                       color="% Ausentes",
                       color_continuous_scale=["#27AE60","#F1C40F","#E74C3C"],
                       text="% Ausentes")
        fig_c.add_hline(y=20, line_dash="dash", line_color="#E74C3C", annotation_text="Alerta 20%")
        fig_c.update_traces(texttemplate="%{text}%", textposition="outside")
        fig_c.update_layout(height=320, plot_bgcolor="white",
                            paper_bgcolor="rgba(0,0,0,0)",
                            margin=dict(t=20,b=10), coloraxis_showscale=False,
                            yaxis_title="% do time ausente")
        st.plotly_chart(fig_c, use_container_width=True)

    st.markdown('<div class="section-title">Ranking — Dias de Férias Agendados por Zupper</div>', unsafe_allow_html=True)
    if "Zupper" in aus.columns and total_col in aus.columns:
        rank = aus[["Zupper","Projeto","Perfil", total_col]].copy()
        rank[total_col] = pd.to_numeric(rank[total_col], errors="coerce").fillna(0)
        rank = rank.sort_values(total_col, ascending=True)
        fig_r = px.bar(rank, x=total_col, y="Zupper", orientation="h",
                       color=total_col,
                       color_continuous_scale=["#BDD7EE","#2C1A1A"],
                       hover_data=["Projeto","Perfil"],
                       text=total_col,
                       labels={total_col: "Dias Agendados"})
        fig_r.update_traces(textposition="outside")
        fig_r.update_layout(height=max(300, len(rank)*28),
                            plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                            margin=dict(t=10,b=10,r=40), coloraxis_showscale=False)
        st.plotly_chart(fig_r, use_container_width=True)

    st.markdown('<div class="section-title">⚠️ Sobreposição de Ausências por Projeto</div>', unsafe_allow_html=True)
    alertas = []
    if "Projeto" in aus.columns:
        for i, mes in enumerate(MESES_NOMES):
            offset_d = 6 + i * 3 + 2
            if offset_d < len(cols):
                col_d = cols[offset_d]
                grp   = aus[aus["Projeto"].astype(str).str.strip() != ""].copy()
                grp["dias_num"] = pd.to_numeric(grp[col_d], errors="coerce").fillna(0)
                ausentes = grp[grp["dias_num"] > 0].groupby("Projeto")["Zupper"].apply(list).reset_index()
                ausentes.columns = ["Projeto","Ausentes"]
                ausentes["Qtd"]  = ausentes["Ausentes"].apply(len)
                criticos = ausentes[ausentes["Qtd"] >= 2]
                for _, row in criticos.iterrows():
                    alertas.append({
                        "Mês": mes,
                        "Projeto": row["Projeto"],
                        "Pessoas Ausentes": row["Qtd"],
                        "Quem": ", ".join(row["Ausentes"])
                    })

        if alertas:
            df_alert = pd.DataFrame(alertas)
            def color_alert(val):
                if isinstance(val, int) and val >= 3: return "background-color:#FFCDD2"
                if isinstance(val, int) and val == 2: return "background-color:#FFF9C4"
                return ""
            styled_alert = df_alert.style.map(color_alert, subset=["Pessoas Ausentes"])
            st.dataframe(styled_alert, use_container_width=True, hide_index=True)
        else:
            st.success("Nenhuma sobreposição crítica de ausências encontrada.")


# ════════════════════════════════════════════════════════════════
# 🚀 ACOMPANHAMENTO DE DEPLOY
# ════════════════════════════════════════════════════════════════
elif pagina == "🚀 Acompanhamento de Deploy":
    st.markdown("# Acompanhamento de Deploy")

    deploy_data = load_deploy()

    if not deploy_data["ok"] or not deploy_data["abas"]:
        st.error(f"❌ Erro ao carregar planilha de deploy: {deploy_data.get('erro', 'Arquivo não encontrado')}")
        st.info("Verifique se o arquivo 'template_deploy_zup.xlsx' está no repositório GitHub.")
        st.stop()

    abas = deploy_data["abas"]

    def fmt_aba(a):
        return f"{a[:2]}/{a[2:4]}/{a[4:]}  {'⬅ atual' if a == abas[0] else ''}"

    ciclo_sel = st.selectbox("Ciclo de deploy", abas, format_func=fmt_aba, index=0)
    df = deploy_data["sheets"][ciclo_sel].copy()

    def tem(val):
        return bool(str(val).strip()) and str(val).strip().lower() not in ["nan","none",""]

    def inferir_status(row):
        pr    = str(row.get("PR",    "")).strip()
        merge = str(row.get("MERGE", "")).strip()
        tag   = str(row.get("TAG",   "")).strip()
        if pr and merge and tag:
            return "Concluído"
        elif pr and merge:
            return "Aguardando TAG"
        elif pr:
            return "Aguardando Merge"
        else:
            return "Pendente"

    df["__status"]       = df.apply(inferir_status, axis=1)
    df["__tem_pr"]       = df["PR"].apply(tem)
    df["__tem_merge"]    = df["MERGE"].apply(tem)
    df["__tem_tag"]      = df["TAG"].apply(tem)
    df["__tem_rollback"] = df["Plano de Rollback"].apply(tem)
    df["__tem_config"]   = df["Configurações"].apply(tem)
    df["__tem_artefato"] = df["Artefatos"].apply(tem)

    total        = len(df)
    concluidos   = len(df[df["__status"] == "Concluído"])
    ag_merge     = len(df[df["__status"] == "Aguardando Merge"])
    ag_tag       = len(df[df["__status"] == "Aguardando TAG"])
    pendentes    = len(df[df["__status"] == "Pendente"])
    prontos      = len(df[df["__tem_pr"] & df["__tem_merge"] & df["__tem_tag"] & df["__tem_rollback"]])

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    k1.markdown(card("Total de Bugs",     total),               unsafe_allow_html=True)
    k2.markdown(card("Concluídos",        concluidos, "green"), unsafe_allow_html=True)
    k3.markdown(card("Ag. Merge",         ag_merge,   "orange"),unsafe_allow_html=True)
    k4.markdown(card("Ag. TAG",           ag_tag,     "orange"),unsafe_allow_html=True)
    k5.markdown(card("Pendentes",         pendentes,  "red"),   unsafe_allow_html=True)
    k6.markdown(card("Prontos p/ Deploy", prontos,    "blue"),  unsafe_allow_html=True)

    st.markdown("---")

    col_check, col_resp = st.columns([1,1])

    with col_check:
        st.markdown('<div class="section-title">Checklist de prontidão</div>', unsafe_allow_html=True)

        def check_line(label, qtd, total_bugs):
            pct  = qtd / total_bugs * 100 if total_bugs > 0 else 0
            cor  = "#27AE60" if pct == 100 else "#F1C40F" if pct >= 50 else "#E74C3C"
            icone = "✅" if pct == 100 else "⚠️" if pct >= 50 else "❌"
            st.markdown(
                f"""<div style="display:flex;align-items:center;gap:10px;margin-bottom:6px;">
                    <span style="font-size:16px">{icone}</span>
                    <span style="flex:1;font-size:13px">{label}</span>
                    <span style="font-size:13px;font-weight:500;color:{cor}">{qtd}/{total_bugs}</span>
                </div>""",
                unsafe_allow_html=True
            )
            st.progress(int(pct) / 100)

        check_line("PR aberto",         df["__tem_pr"].sum(),       total)
        check_line("Merge realizado",   df["__tem_merge"].sum(),    total)
        check_line("TAG gerada",        df["__tem_tag"].sum(),      total)
        check_line("Plano de Rollback", df["__tem_rollback"].sum(), total)
        check_line("Configurações",     df["__tem_config"].sum(),   total)
        check_line("Artefatos",         df["__tem_artefato"].sum(), total)

    with col_resp:
        st.markdown('<div class="section-title">Bugs por responsável</div>', unsafe_allow_html=True)
        if "RESPONSÁVEL" in df.columns:
            resp_cnt = df["RESPONSÁVEL"].value_counts().reset_index()
            resp_cnt.columns = ["Responsável","Qtd"]
            fig_resp = px.bar(resp_cnt, x="Qtd", y="Responsável", orientation="h",
                              color_discrete_sequence=["#E05A2B"], text="Qtd")
            fig_resp.update_traces(textposition="outside")
            fig_resp.update_layout(height=280, plot_bgcolor="white",
                                   paper_bgcolor="rgba(0,0,0,0)",
                                   margin=dict(t=10,b=10,l=10,r=40),
                                   xaxis_title="", yaxis_title="")
            st.plotly_chart(fig_resp, use_container_width=True)

    st.markdown("---")

    st.markdown('<div class="section-title">Bugs do ciclo</div>', unsafe_allow_html=True)

    status_cores = {
        "Concluído":       "#C8E6C9",
        "Aguardando TAG":  "#FFF9C4",
        "Aguardando Merge":"#FFE0B2",
        "Pendente":        "#FFCDD2",
    }

    show_cols = [c for c in ["RESPONSÁVEL","BUG","Descrição","__status",
                              "PR","MERGE","TAG","Plano de Rollback",
                              "Configurações","Artefatos"] if c in df.columns]

    df_show = df[show_cols].rename(columns={"__status": "Status"})

    def color_status(val):
        return f"background-color:{status_cores.get(val,'')}"

    styled = df_show.style.map(color_status, subset=["Status"])
    st.dataframe(styled, use_container_width=True, height=420, hide_index=True)

    st.markdown("---")

    st.markdown('<div class="section-title">Alertas automáticos</div>', unsafe_allow_html=True)

    sem_rb_lista  = df[~df["__tem_rollback"]]["BUG"].tolist()
    sem_tag_lista = df[df["__tem_merge"] & ~df["__tem_tag"]]["BUG"].tolist()
    prontos_lista = df[df["__tem_pr"] & df["__tem_merge"] & df["__tem_tag"] & df["__tem_rollback"]]["BUG"].tolist()

    al1, al2, al3 = st.columns(3)

    with al1:
        if sem_rb_lista:
            st.error(f"**⛔ Sem plano de rollback ({len(sem_rb_lista)})**\n\n" +
                     "\n".join(f"- {b}" for b in sem_rb_lista))
        else:
            st.success("✅ Todos com rollback documentado")

    with al2:
        if sem_tag_lista:
            st.warning(f"**⚠️ Merge feito, TAG pendente ({len(sem_tag_lista)})**\n\n" +
                       "\n".join(f"- {b}" for b in sem_tag_lista))
        else:
            st.success("✅ Todos com TAG gerada")

    with al3:
        if prontos_lista:
            st.info(f"**🚀 Prontos para deploy ({len(prontos_lista)})**\n\n" +
                    "\n".join(f"- {b}" for b in prontos_lista))
        else:
            st.warning("Nenhum bug 100% pronto ainda")

    if len(abas) > 1:
        st.markdown("---")
        st.markdown('<div class="section-title">Histórico de ciclos</div>', unsafe_allow_html=True)

        hist = []
        for aba in abas:
            d = deploy_data["sheets"][aba].copy()
            if d.empty:
                continue
            d["__status"] = d.apply(inferir_status, axis=1)
            hist.append({
                "Ciclo":      f"{aba[:2]}/{aba[2:4]}/{aba[4:]}",
                "Total":      len(d),
                "Concluídos": len(d[d["__status"] == "Concluído"]),
                "Pendentes":  len(d[d["__status"] == "Pendente"]),
            })

        df_hist = pd.DataFrame(hist)
        fig_hist = px.bar(df_hist, x="Ciclo", y=["Concluídos","Pendentes"],
                          barmode="group",
                          color_discrete_sequence=["#27AE60","#E74C3C"],
                          labels={"value":"Bugs","variable":""})
        fig_hist.update_layout(height=280, plot_bgcolor="white",
                               paper_bgcolor="rgba(0,0,0,0)",
                               margin=dict(t=10,b=10),
                               legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig_hist, use_container_width=True)


# ════════════════════════════════════════════════════════════════
# 🏢 ESTRUTURA DAS SQUADS
# ════════════════════════════════════════════════════════════════
elif pagina == "🏢 Estrutura das Squads":
    st.markdown("# Estrutura das Squads · Flex 2026")
    st.caption("Visão completa da organização dos times por líder de capítulo e squad.")

    # ── Dados da estrutura ────────────────────────────────────────
    SQUADS = {
        "carlos": {
            "titulo": "Carlos (NP)",
            "squads": {
                "Salesforce": {
                    "cor": "#0070D2",
                    "membros": [
                        ("Edson Vendramini","LT"),("Luis Sales","LT"),("Rafael Pasquim","LT"),
                        ("Juliano Stercket","SM"),("William Carmo","PO"),("Marcleide Sousa","QA"),
                        ("Leandro Moral","Dev"),("José Rodrigues","Dev"),("Henrique Tosta","Dev"),
                        ("Nadia Lima","Dev"),("Danilo Silva","Dev"),("Fernando Assunção","LT"),
                    ]
                },
            }
        },
        "victor": {
            "titulo": "Victor (Banda Larga e Portal)",
            "squads": {
                "Silver":      {"cor":"#6B7280","membros":[("Tiago Carpanese","LT"),("Diogo Arantes","SM"),("Sônia Villanova","PO"),("Lucas Paglia Grisa","Back"),("Murilo Oliari Ribeiro","Back"),("Flavio de Paula Faria","Back"),("Ricardo Lima","Back"),("Danilo Valente dos Santos","Back"),("Herbert Luiz Soares","Back"),("Vivian Chiodo Dias","iOS"),("Edmundo Schultz Neto","And"),("Pedro Henrique Borges","iOS"),("Guilherme Jose","And"),("Maxwell de Sousa Martins","QA"),("Jéssica Euzébio Rocha","QA"),("Willian Moya","AT")]},
                "Golden":      {"cor":"#B8860B","membros":[("Tiago Carpanese","TL"),("Larissa Trinchão","SM"),("Felipe Colem","PO"),("Filipe Alves Pinheiro","Back"),("Muriel Magno Teles","Back"),("Breno Silva","Back"),("Gerson Arbigaus","iOS"),("Erick Teixeira","iOS"),("Luan de Souza","iOS"),("Alexandre Mahmud","And"),("Juan Vicente","And"),("Bianca Bueno","QA"),("Victor Vital","QA"),("David Gomes","AT"),("Bruno Ferronato","Dev")]},
                "Purple":      {"cor":"#6B21A8","membros":[("Vinicius Simone","SM"),("Alexandre Marques","PO"),("Fabio de Souza","Back"),("Paulo Gonçalves","Back"),("Carlos Silva","Back"),("Sylas Eckart","Web"),("Barbara Perina","QA"),("Silvana de Souza","QA"),("Marcelo Moreira","QA"),("Yuri Duarte Alves","QA"),("Willian Moya","Esp")]},
                "Red":         {"cor":"#DC2626","membros":[("Vinicius Vense","LT"),("Allan Ribeiro","SM"),("Henrique Franzão","PO"),("Rodolfo Souza","Back"),("Jaqueline Moreno","Back"),("Vitor Lopes","Back"),("Luan Almeida","iOS"),("Italo Bianchini","iOS"),("Rui Barbosa","And"),("Raphael Silva","And"),("Kim","iOS")]},
                "Yellow (BL)": {"cor":"#CA8A04","membros":[("Rafael Gomes","LT"),("Débora Ferreira","SM"),("PV Freitas","PO"),("Djair Soares Pereira","Back"),("Maria de Fatima","Back"),("João Henrique De Sá","Back"),("Lucas Terra","Back"),("Renan Veloso Silva","iOS"),("Raline Maria da Silva","iOS"),("Francisco Pereira","And"),("Wiliam Trancoso","And"),("Lucas Christopher","And"),("Brenda De Souza Garcia","QA"),("Diego Soares Santana","QA"),("Wagner","QA"),("Amabily de Oliveira","QA"),("George Tassiano","Web"),("Lucas Limoni","Web"),("Gabrielle Martins","Web"),("Michel Rodrigues","Web"),("Lucas Augusto Caetano","QA"),("Tiago Gomes","QA"),("Danielle de Brito","QA")]},
                "Blue":        {"cor":"#1D4ED8","membros":[("Marcionei Bizerra","PO"),("Rejane da Silva","PO"),("Carol Masaki","SM"),("Gabriel Filipe","Back"),("Marcelo Littig","Back"),("Ronaldo Souza Cutrim","CSP"),("Roseane Costa","Back"),("Jeferson Fernandes","And"),("Pedro Henrique de Oliveira","iOS"),("Caroline Gomes","QA"),("Eduardo Rodrigues","QA"),("Marcus Lucena","LT"),("Aparicio Neto","Web"),("Dieison Silva","Web"),("Rosenilton Reis","Web"),("Wellyngton Matheus","Web")]},
            }
        },
        "michele": {
            "titulo": "Michele",
            "squads": {
                "White": {"cor":"#374151","membros":[("Hugo Celestino","LT"),("Monaliza Felipe","SM"),("Julio Ferreira Junior","PO"),("Vitoria Silva Cardoso","Back"),("Gabriel Severino","Back"),("Henio De Alcantara Junior","Back"),("Arthur Amestrete","Back"),("João Pedro Franco","iOS"),("Renan Maganha","iOS"),("Gabriel Hernandes","And"),("Mauricio de Souza Martins","And"),("Luiz Gustavo Fleria","QA"),("Felipe Falsetta","QA"),("Brenno Leite","QA"),("Carla Gomes","QA")]},
                "Green": {"cor":"#15803D","membros":[("Thasso / Hugo","LT"),("Yago Taveiros Ferreira","LT"),("Rafael Lusivo","SM"),("Márcia Beatriz","PO"),("Anderson Barreiro","Back"),("Jefferson","Back"),("Beatriz Martins","Back"),("Pablo Henrique Castro","Back"),("Michel Peixoto","Back"),("Gerson","QA Back"),("Lucas Dittrich","QA Back"),("Matheus Fusco","iOS"),("Thais Conde","iOS"),("Matheus Santana","And"),("Ronnyery Barbosa","And"),("Letícia Vale","QA Front"),("Dreice Cousino","QA Front"),("Filipe Malta","QA Front"),("Daniel Ferreira","Back"),("Daniel Ramos","Back")]},
            }
        },
        "joao": {
            "titulo": "João Junior",
            "squads": {
                "Black": {"cor":"#111827","membros":[("Nayara Gaspar","LT"),("Lorena Melo","SM"),("Elizeu Rocha","PO"),("Eduardo Morgon","Back"),("Eduardo Rodrigues Delfino","Back"),("Felipe Ferreira Rezende","Back"),("Jefferson Santos Oliveira","Back"),("Raphaela Rosa","Back"),("Marcilio Zanatta","Back"),("José Vitor de A. Coelho","Back"),("Francisco Clewerton Pereira Roque","Back"),("Renato Filho","iOS"),("Fabio Franca","iOS"),("Joao Eudes","And"),("Augusto Favretto","And"),("Marcos Vinicius Macedo de Menezes","Web"),("Thiago Antonio Rodrigues","QA"),("Jonatas Micael Silva Peixoto","QA"),("Arthur Ayres","QA"),("Pablo Fleria","QA"),("Rhaniel Farias","DEV"),("Evelyn Sthefany","QA")]},
            }
        },
        "raj": {
            "titulo": "Raj",
            "squads": {
                "Orange / TA":  {"cor":"#EA580C","membros":[("Alysson Pereira","LT"),("Alexsandra Correa","QA"),("Victor Guerra","QA")]},
                "Evoluções":    {"cor":"#0891B2","membros":[("Thasso Araujo","LT"),("Alexsandra Correa da Rosa","QA"),("Vinicius Freitas","Back NP"),("Nilton Mitsuharu Sugawara","Back"),("Gillian Mendes da Costa","Back"),("Luiz Henrique Bortolini","Back")]},
                "Segurança":    {"cor":"#7C3AED","membros":[("Lucas Gontijo de Souza","Back Spec"),("Edigar Ferreira Junior","Back"),("Vaga aberta","Back Spec"),("Gillian Mendes da Costa","Back"),("Jessamine Silva Almeida Bueno","QA")]},
            }
        },
    }

    # ── Controles ─────────────────────────────────────────────────
    col_v1, col_v2 = st.columns([3, 1])
    with col_v1:
        busca = st.text_input("🔍 Buscar membro", placeholder="Digite parte do nome...")
    with col_v2:
        visao = st.selectbox("Visão", ["Interativa (com filtros)", "Completa (todos os membros)"])

    c1, c2, c3 = st.columns(3)
    with c1:
        lider_sel = st.selectbox("Líder de capítulo", ["Todos"] + [v["titulo"] for v in SQUADS.values()])
    with c2:
        cargo_opts = ["Todos","LT","SM","PO","Back","iOS","And","QA","Web","Dev","AT"]
        cargo_sel  = st.selectbox("Cargo / Perfil", cargo_opts)
    with c3:
        todos_squads = []
        for cap_data in SQUADS.values():
            todos_squads.extend(cap_data["squads"].keys())
        squad_sel = st.selectbox("Squad", ["Todas"] + sorted(todos_squads))

    st.markdown("---")

    # ── KPIs dinâmicos ────────────────────────────────────────────
    total_membros  = sum(len(sq["membros"]) for cap in SQUADS.values() for sq in cap["squads"].values())
    total_squads   = sum(len(cap["squads"]) for cap in SQUADS.values())
    total_lideres  = len(SQUADS)
    total_lts      = sum(1 for cap in SQUADS.values() for sq in cap["squads"].values() for _,c in sq["membros"] if c == "LT")

    mk1,mk2,mk3,mk4 = st.columns(4)
    mk1.markdown(card("Total de membros", total_membros, "orange"), unsafe_allow_html=True)
    mk2.markdown(card("Squads ativos",    total_squads,  "blue"),   unsafe_allow_html=True)
    mk3.markdown(card("Líderes de capítulo", total_lideres),        unsafe_allow_html=True)
    mk4.markdown(card("Líderes Técnicos (LT)", total_lts, "green"), unsafe_allow_html=True)

    st.markdown("---")

    # ── Hierarquia topo ───────────────────────────────────────────
    st.markdown("""
    <div style="text-align:center;margin-bottom:8px">
      <span class="squad-dir">Norberto Imamura<small>Delivery</small></span>
      <span class="squad-dir">Rafael Moreira<small>Operações</small></span>
    </div>
    <div style="text-align:center;margin-bottom:4px">
      <span class="squad-coord">Katia Santos</span>
    </div>
    <div style="text-align:center;margin-bottom:12px">
      <span class="squad-coord">Heleno Araújo<small>Backend</small></span>
      <span class="squad-coord">Pri Barone<small>Frontend</small></span>
    </div>
    """, unsafe_allow_html=True)

    # ── Renderiza capítulos ───────────────────────────────────────
    busca_lower = busca.strip().lower()

    for cap_key, cap_data in SQUADS.items():
        # Filtro por líder
        if lider_sel != "Todos" and cap_data["titulo"] != lider_sel:
            continue

        # Verifica se há algo para mostrar
        squads_visiveis = {}
        for sq_nome, sq_data in cap_data["squads"].items():
            if squad_sel != "Todas" and sq_nome != squad_sel:
                continue
            membros_filtrados = []
            for nome, cargo in sq_data["membros"]:
                if cargo_sel != "Todos" and cargo_sel.lower() not in cargo.lower():
                    continue
                if busca_lower and busca_lower not in nome.lower():
                    continue
                membros_filtrados.append((nome, cargo))
            if membros_filtrados or visao == "Completa (todos os membros)":
                squads_visiveis[sq_nome] = (sq_data["cor"], membros_filtrados if membros_filtrados else sq_data["membros"])

        if not squads_visiveis:
            continue

        total_cap = sum(len(m) for _, m in squads_visiveis.values())

        with st.expander(f"**{cap_data['titulo']}** — {len(squads_visiveis)} squad(s) · {total_cap} membros", expanded=True):
            cols_squads = st.columns(min(len(squads_visiveis), 4))
            for idx, (sq_nome, (cor, membros_lista)) in enumerate(squads_visiveis.items()):
                col_idx = idx % min(len(squads_visiveis), 4)
                with cols_squads[col_idx]:
                    st.markdown(
                        f'<div class="sq-pill" style="background:{cor}">{sq_nome} · {len(membros_lista)}</div>',
                        unsafe_allow_html=True
                    )
                    tags_html = ""
                    for nome, cargo in membros_lista:
                        hl = "hl" if busca_lower and busca_lower in nome.lower() else ""
                        tags_html += f'<span class="membro-tag {hl}">{nome} <b>{cargo}</b></span>'
                    st.markdown(tags_html, unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)

    # ── Gráfico distribuição por cargo ───────────────────────────
    st.markdown("---")
    st.markdown('<div class="section-title">Distribuição por cargo técnico</div>', unsafe_allow_html=True)

    cargo_count: dict = {}
    for cap in SQUADS.values():
        for sq in cap["squads"].values():
            for _, cargo in sq["membros"]:
                cargo_count[cargo] = cargo_count.get(cargo, 0) + 1

    df_cargos = pd.DataFrame(sorted(cargo_count.items(), key=lambda x: -x[1]), columns=["Cargo","Membros"])
    fig_cargos = px.bar(
        df_cargos, x="Membros", y="Cargo", orientation="h",
        color="Membros", color_continuous_scale=["#BDD7EE","#E05A2B"],
        text="Membros"
    )
    fig_cargos.update_traces(textposition="outside")
    fig_cargos.update_layout(
        height=max(300, len(df_cargos)*32),
        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=10,b=10,r=40), coloraxis_showscale=False
    )
    st.plotly_chart(fig_cargos, use_container_width=True)
