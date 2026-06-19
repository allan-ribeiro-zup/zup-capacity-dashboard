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

# ── Google Drive ──────────────────────────────────────────────────
GDRIVE_FILE_ID = "1IFhx6IEksLFI0AzVSZHDroaw8GKnw7yA"
GDRIVE_URL     = f"https://drive.google.com/uc?export=download&id={GDRIVE_FILE_ID}"

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

@st.cache_data(show_spinner="Carregando dados do Google Drive...", ttl=300)
def load():
    try:
        resp = requests.get(GDRIVE_URL, timeout=30)
        resp.raise_for_status()
        content = io.BytesIO(resp.content)
        xls   = pd.ExcelFile(content)
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
        return {
            "membros":    get("Membros"),
            "ausencias":  get("Ausências"),
            "cap":        get("Capacidade por Sprint"),
            "roadmap":    get("Roadmap"),
            "atividades": get("Atividades por Release"),
            "ok":         True,
        }
    except Exception as e:
        return {"ok": False, "erro": str(e)}

# ── Carrega dados ─────────────────────────────────────────────────
with st.spinner("Conectando ao Google Drive..."):
    data = load()

if not data.get("ok"):
    st.error(f"❌ Erro ao carregar planilha: {data.get('erro','')}")
    st.info("Verifique se o arquivo está compartilhado como 'Qualquer pessoa com o link pode visualizar'.")
    st.stop()

roadmap   = data["roadmap"]
ativ      = data["atividades"]
cap       = data["cap"]
membros   = data["membros"]
ausencias = data["ausencias"]

# 🏠 DASHBOARD
# ════════════════════════════════════════════════════════════════
if pagina == "🏠 Dashboard":
    st.markdown("# Dashboard de Capacidade · Gestão de Portfólio")

    # Filtros
    squads_all = sorted(roadmap["Squad"].dropna().unique().tolist()) if "Squad" in roadmap.columns else []
    releases_all = sorted(ativ["Release"].dropna().unique().tolist()) if "Release" in ativ.columns else []

    cf1, cf2 = st.columns([3,1])
    with cf1:
        squad_sel = st.multiselect("Filtrar Squad", squads_all, default=squads_all)
    with cf2:
        rel_sel = st.selectbox("Release", ["Todas"] + releases_all)

    rm = roadmap.copy()
    if squad_sel and "Squad" in rm.columns:
        rm = rm[rm["Squad"].isin(squad_sel)]

    # KPIs
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

    # Donut status
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

    # Horas por Squad (barras)
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

    # Capacidade por Sprint (linha + barra)
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

    # Heatmap plataforma × sprint
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

    # Filtros
    f1,f2,f3,f4 = st.columns(4)
    squads_r   = sorted(roadmap["Squad"].dropna().unique())    if "Squad"      in roadmap.columns else []
    status_r   = sorted(roadmap["Status"].dropna().unique())   if "Status"     in roadmap.columns else []
    plats_r    = sorted(roadmap["Plataforma"].dropna().unique()) if "Plataforma" in roadmap.columns else []

    with f1: sq  = st.multiselect("Squad",      squads_r)
    with f2: st_ = st.multiselect("Status",     status_r)
    with f3: pl  = st.multiselect("Plataforma", plats_r)
    with f4: bsc = st.text_input("🔍 Buscar tarefa")

    rm2 = roadmap.copy()
    if sq  and "Squad"      in rm2.columns: rm2 = rm2[rm2["Squad"].isin(sq)]
    if st_ and "Status"     in rm2.columns: rm2 = rm2[rm2["Status"].isin(st_)]
    if pl  and "Plataforma" in rm2.columns: rm2 = rm2[rm2["Plataforma"].isin(pl)]
    if bsc and "Tarefa"     in rm2.columns: rm2 = rm2[rm2["Tarefa"].str.contains(bsc, case=False, na=False)]

    # KPIs filtrados
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

    # Mini charts
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

    # Heatmap de % por sprint
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

    # Barras de progresso por iniciativa
    st.markdown('<div class="section-title">Status por Iniciativa</div>', unsafe_allow_html=True)
    if "Grande Iniciativa" in df_rel.columns and "Status" in df_rel.columns:
        for _, row in df_rel.iterrows():
            ini    = row.get("Grande Iniciativa","—")
            status = row.get("Status","Não Iniciado")
            squad  = row.get("Squad","")
            # calcula % médio das sprints com valor
            vals = [row.get(s) for s in sp_cols if pd.notna(row.get(s)) and isinstance(row.get(s),(int,float))]
            pct  = max(vals) if vals else 0
            cor  = STATUS_COLORS.get(status,"#BDC3C7")
            col_i, col_p = st.columns([4,6])
            with col_i:
                st.markdown(f"**{ini}** <span style='font-size:11px;color:#888'>· {squad}</span>", unsafe_allow_html=True)
                st.caption(status)
            with col_p:
                st.progress(int(pct)/100, text=f"{int(pct)}%")

    # Tabela
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

    # KPIs plataforma
    plats = sorted(cap_f["Plataforma"].dropna().unique()) if "Plataforma" in cap_f.columns else []
    if plats:
        cols_k = st.columns(len(plats))
        for idx, plat in enumerate(plats):
            df_p   = cap_f[cap_f["Plataforma"]==plat]
            total  = df_p[sp_cols].sum().sum() if sp_cols else 0
            n_mem  = len(membros[membros["Plataforma"]==plat]) if not membros.empty and "Plataforma" in membros.columns else "—"
            with cols_k[idx]:
                st.markdown(f"""<div class="metric-card">
                    <div class="metric-label">{plat}</div>
                    <div class="metric-value orange">{total:.0f}h</div>
                    <div style="font-size:11px;color:#888">{n_mem} membros</div>
                </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # Gráfico capacidade por sprint e plataforma
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

    # Heatmap %
    st.markdown('<div class="section-title">Heatmap de Horas por Plataforma × Sprint</div>', unsafe_allow_html=True)
    if sp_cols and "Plataforma" in cap_f.columns:
        pvt = cap_f.groupby("Plataforma")[sp_cols].sum()
        fig_hm = px.imshow(pvt, text_auto=".0f",
                           color_continuous_scale=["#FFEEEE","#F1C40F","#27AE60"],
                           aspect="auto",
                           labels=dict(x="Sprint", y="Plataforma", color="Horas"))
        fig_hm.update_layout(height=230, margin=dict(t=5,b=5), coloraxis_showscale=False)
        st.plotly_chart(fig_hm, use_container_width=True)

    # Membros por squad
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
    st.markdown("# Calendário de Férias e Ausências")

    if ausencias.empty:
        st.warning("Sem dados na aba 'Ausências'."); st.stop()

    squads_a = sorted(ausencias["Squad"].dropna().unique()) if "Squad" in ausencias.columns else []
    tipos_a  = sorted(ausencias["Tipo"].dropna().unique())  if "Tipo"  in ausencias.columns else []

    fa1, fa2 = st.columns(2)
    with fa1: sq_a = st.multiselect("Squad", squads_a, default=squads_a)
    with fa2: tp_a = st.multiselect("Tipo",  tipos_a,  default=tipos_a)

    aus = ausencias.copy()
    if sq_a and "Squad" in aus.columns: aus = aus[aus["Squad"].isin(sq_a)]
    if tp_a and "Tipo"  in aus.columns: aus = aus[aus["Tipo"].isin(tp_a)]

    total_h = aus["Horas Indisponíveis"].sum() if "Horas Indisponíveis" in aus.columns else 0
    pessoas = aus["Nome"].nunique()             if "Nome"               in aus.columns else 0

    ka,kb,kc = st.columns(3)
    ka.markdown(card("Registros",          len(aus)          ), unsafe_allow_html=True)
    kb.markdown(card("Pessoas c/ Ausência", pessoas, "orange"), unsafe_allow_html=True)
    kc.markdown(card("Horas Indisponíveis", f"{total_h:.0f}h","red"), unsafe_allow_html=True)

    st.markdown("---")

    # Tabela colorida
    st.markdown('<div class="section-title">Registros de Ausência</div>', unsafe_allow_html=True)
    tipo_bg = {"Férias":"#BBDEFB","Folga":"#FFF9C4","Licença":"#FFCDD2"}

    def color_tipo(v):
        return f"background-color:{tipo_bg.get(v,'')}"

    show_a = [c for c in ["Nome","Squad","Tipo","Data Início","Data Fim",
                           "Sprint(s) Afetado(s)","Horas Indisponíveis"] if c in aus.columns]
    if "Tipo" in aus.columns:
        styled_a = aus[show_a].style.map(color_tipo, subset=["Tipo"])
        st.dataframe(styled_a, use_container_width=True, hide_index=True)
    else:
        st.dataframe(aus[show_a], use_container_width=True, hide_index=True)

    # Impacto por sprint
    st.markdown('<div class="section-title">Impacto por Sprint</div>', unsafe_allow_html=True)
    if "Sprint(s) Afetado(s)" in aus.columns and "Horas Indisponíveis" in aus.columns:
        sp_imp = aus.groupby("Sprint(s) Afetado(s)")["Horas Indisponíveis"].sum().reset_index()
        fig_a = px.bar(sp_imp, x="Sprint(s) Afetado(s)", y="Horas Indisponíveis",
                       color_discrete_sequence=["#E05A2B"],
                       labels={"Sprint(s) Afetado(s)":"Sprint","Horas Indisponíveis":"Horas Indisponíveis"})
        fig_a.update_layout(height=280, plot_bgcolor="white",
                            paper_bgcolor="rgba(0,0,0,0)", margin=dict(t=10,b=10))
        st.plotly_chart(fig_a, use_container_width=True)

    # Por tipo
    st.markdown('<div class="section-title">Distribuição por Tipo</div>', unsafe_allow_html=True)
    if "Tipo" in aus.columns and not aus.empty:
        tp_cnt = aus["Tipo"].value_counts().reset_index()
        tp_cnt.columns = ["Tipo","Count"]
        fig_t = px.pie(tp_cnt, values="Count", names="Tipo", hole=0.45,
                       color="Tipo",
                       color_discrete_map={"Férias":"#BBDEFB","Folga":"#FFF9C4","Licença":"#FFCDD2"})
        fig_t.update_layout(height=260, margin=dict(t=10,b=10,l=10,r=80))
        st.plotly_chart(fig_t, use_container_width=True)
