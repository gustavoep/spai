import streamlit as st
import pandas as pd
import plotly.express as px
import os
import unicodedata
from datetime import datetime, timezone, timedelta

# Caminhos exatos no GitHub / Repositório
CAMINHO_SEGMENTOS = "data/Fato_Segmentos.xlsx"
CAMINHO_VERTICAIS = "data/Fato_Verticais.xlsx"
CAMINHO_CTI = "assets/CTI.png"
CAMINHO_LOGO_SCTI = "assets/SCTI.png"
CAMINHO_LOGO_HORIZONTAL = "assets/Logotipo Horizontal SPAI.png"
CAMINHO_LOGO_JAPI = "assets/Logo_JAPI.png"

# Fuso horário oficial de Brasília (UTC-3)
FUSO_BRASILIA = timezone(timedelta(hours=-3))

# Função para normalizar strings (remove acentos, maiúsculas e espaços extras)
def normalizar_texto(txt):
    if not isinstance(txt, str):
        txt = str(txt) if pd.notna(txt) else ""
    txt = txt.strip().lower()
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('utf-8')
    return txt

# Função para ordenar a lista de categorias priorizando Parque Tecnológico, Incubadoras, Centros de Inovação e Centros de Conexão
def ordenar_categorias_prioritarias(lista_cats):
    ordem_prioritaria = [
        "parque tecnologico", "parques tecnologicos",
        "incubadora", "incubadoras",
        "centro de inovacao", "centros de inovacao",
        "centro de conexao", "centros de conexao"
    ]
    def prioridade_chave(cat):
        c_norm = normalizar_texto(cat)
        for item in ordem_prioritaria:
            if item in c_norm or c_norm in item:
                if "parque" in item: return 0
                if "incubadora" in item: return 1
                if "centro de inovacao" in item or "centros de inovacao" in item: return 2
                if "centro de conexao" in item or "centros de conexao" in item: return 3
        return 999
    return sorted(lista_cats, key=lambda c: (prioridade_chave(c), c))

# Configuração da Página (Primeiro comando Streamlit)
st.set_page_config(
    page_title="SPAI - Ambientes de Inovação SP",
    page_icon=CAMINHO_CTI if os.path.exists(CAMINHO_CTI) else "🏛️",
    layout="wide"
)

# URL ao vivo da planilha do Google Sheets (Base Geral SPAI)
GOOGLE_SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1yXVrdrAZW9QCYHIfp2MsyIa2iFsVAhGZYxnOst_UG_M/export?format=csv&gid=1170767452"

# Mapeamento de Coordenadas dos Municípios Paulistas
COORDENADAS_SP = {
    "São Paulo": (-23.5505, -46.6333), "Campinas": (-22.9099, -47.0626), "São José dos Campos": (-23.1791, -45.8872),
    "Sorocaba": (-23.5015, -47.4526), "Ribeirão Preto": (-21.1704, -47.8103), "Santos": (-23.9608, -46.3339),
    "Piracicaba": (-22.7253, -47.6492), "São Carlos": (-22.0175, -47.8908), "Guarulhos": (-23.4628, -46.5333),
    "Bebedouro": (-20.9494, -48.4792), "Leme": (-22.1864, -47.3878), "Jundiaí": (-23.1857, -46.8892),
    "Botucatu": (-22.8858, -48.4450), "Marília": (-22.2139, -49.9458), "Indaiatuba": (-23.0903, -47.2181),
    "Santo André": (-23.6639, -46.5383), "São José do Rio Preto": (-20.8114, -49.3758), "Presidente Prudente": (-22.1256, -51.3889),
    "Americana": (-22.7392, -47.3314), "Gavião Peixoto": (-21.8383, -48.4975), "Rio Claro": (-22.4147, -47.5614),
    "Guarujá": (-23.9931, -46.2564), "Lins": (-21.6786, -49.7428), "Franca": (-20.5386, -47.4008),
    "Registro": (-24.4881, -47.8436), "Guaratinguetá": (-22.8164, -45.1925), "Itapeva": (-23.9822, -48.8764),
    "Bauru": (-22.3147, -49.0606), "Ilha Solteira": (-20.3858, -51.3414), "Assis": (-22.6608, -50.4158),
    "Pompéia": (-22.1086, -50.1736), "Jaboticabal": (-21.2553, -48.3228), "Araraquara": (-21.7944, -48.1758),
    "Tarumã": (-22.7461, -50.5819), "Barretos": (-20.5572, -48.5678), "Atibaia": (-23.1169, -46.5503),
    "São Caetano do Sul": (-23.6228, -46.5544), "Dracena": (-21.4839, -51.5331), "Garça": (-22.2125, -49.6547),
    "Olímpia": (-20.7381, -48.9142), "São Sebastião": (-23.7600, -45.4100), "Araçatuba": (-21.2089, -50.4328)
}

# Estilização CSS Personalizada - Neutra, Polida e com Destaque para Filtros
st.markdown("""
    <style>
    .block-container { padding-top: 1rem; padding-bottom: 0rem; }

    /* Card de Destaque Executivo - Total de Membros */
    .kpi-box-membros {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        color: white;
        border-radius: 10px;
        padding: 16px 12px;
        text-align: center;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.15);
        border-left: 5px solid #0a9396;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .kpi-val-membros {
        font-size: 38px;
        font-weight: 800;
        line-height: 1;
        color: #ffffff;
        letter-spacing: -0.5px;
    }
    .kpi-lbl-membros {
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-top: 6px;
        color: #94d2bd;
    }

    /* Cards Secundários do Topo */
    .kpi-box-top {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-top: 4px solid #005f73;
        border-radius: 10px;
        padding: 14px 10px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .kpi-val-top { font-size: 32px; font-weight: 800; line-height: 1; color: #005f73; }
    .kpi-lbl-top { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; margin-top: 6px; color: #475569; }

    /* Cards da Segunda Linha (Categorias) */
    .kpi-box-sec {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-top: 3px solid #0a9396;
        border-radius: 8px;
        padding: 10px 6px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .kpi-val-sec { font-size: 24px; font-weight: 800; color: #0f172a; line-height: 1; }
    .kpi-lbl-sec { font-size: 10px; font-weight: 700; color: #475569; text-transform: uppercase; margin-top: 4px; line-height: 1.2; }
    .kpi-sub-sec { font-size: 9px; color: #64748b; font-style: italic; margin-top: 2px; }

    /* Card Destaque - Filtro de Tipo de Membro */
    .filter-card-tipo {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border: 1px solid #cbd5e1;
        border-left: 6px solid #005f73;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 16px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    }
    .filter-card-hdr {
        font-size: 14px;
        font-weight: 800;
        color: #0f172a;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-bottom: 2px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .filter-card-sub {
        font-size: 12px;
        color: #64748b;
        margin-bottom: 10px;
    }

    div[data-testid="stPills"] button {
        border-radius: 8px !important;
        font-weight: 700 !important;
        font-size: 13px !important;
        padding: 8px 18px !important;
        border: 1px solid #94a3b8 !important;
        transition: all 0.2s ease-in-out !important;
    }

    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .metric-val { font-size: 24px; font-weight: bold; color: #005f73; }
    .metric-lbl { font-size: 12px; color: #64748b; font-weight: 600; }

    .concept-card {
        background-color: #f8fafc;
        border-left: 4px solid #005f73;
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 15px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    .concept-title {
        font-size: 15px;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 6px;
    }
    .concept-text {
        font-size: 13px;
        color: #334155;
        line-height: 1.5;
    }
    </style>
""", unsafe_allow_html=True)

# Função de carregamento com cache da Base Geral
@st.cache_data(ttl=120)
def carregar_base_geral():
    df = pd.read_csv(GOOGLE_SHEET_CSV_URL)
    cidade_col = df.iloc[:, 3].astype(str).str.strip() if df.shape[1] > 3 else pd.Series()
    df["CIDADE_CLEAN"] = cidade_col
    df["Lat"] = df["CIDADE_CLEAN"].map(lambda c: COORDENADAS_SP.get(c, (None, None))[0])
    df["Lon"] = df["CIDADE_CLEAN"].map(lambda c: COORDENADAS_SP.get(c, (None, None))[1])
    return df

# Função para carregar arquivo Excel do caminho exato
def carregar_excel_fato(caminho_exato):
    if os.path.exists(caminho_exato):
        try:
            return pd.read_excel(caminho_exato)
        except Exception:
            pass
    return None

df_spai = carregar_base_geral()
df_segmentos = carregar_excel_fato(CAMINHO_SEGMENTOS)
df_verticais = carregar_excel_fato(CAMINHO_VERTICAIS)

# Função para consultar Verticais e Segmentos de um ambiente específico
def buscar_detalhes_ambiente(nome_ambiente):
    key = normalizar_texto(nome_ambiente)
    verts, segs = [], []
    
    if df_verticais is not None and not df_verticais.empty:
        col_n = None
        col_a = None
        for c in df_verticais.columns:
            if any(p in str(c).lower() for p in ["nome", "ambiente", "instituição", "instituicao"]):
                col_n = c
            if any(p in str(c).lower() for p in ["vertical", "segmento", "área", "area", "atuação", "atuacao"]):
                col_a = c
        col_n = col_n or df_verticais.columns[0]
        col_a = col_a or df_verticais.columns[-1]
        
        m = df_verticais[df_verticais[col_n].astype(str).apply(normalizar_texto) == key]
        if not m.empty:
            verts = sorted(m[col_a].dropna().astype(str).str.strip().unique())

    if df_segmentos is not None and not df_segmentos.empty:
        col_n = None
        col_a = None
        for c in df_segmentos.columns:
            if any(p in str(c).lower() for p in ["nome", "ambiente", "instituição", "instituicao"]):
                col_n = c
            if any(p in str(c).lower() for p in ["vertical", "segmento", "área", "area", "atuação", "atuacao"]):
                col_a = c
        col_n = col_n or df_segmentos.columns[0]
        col_a = col_a or df_segmentos.columns[-1]
        
        m = df_segmentos[df_segmentos[col_n].astype(str).apply(normalizar_texto) == key]
        if not m.empty:
            segs = sorted(m[col_a].dropna().astype(str).str.strip().unique())

    return verts, segs

# Modal Popup com as informações detalhadas do Ambiente
if hasattr(st, "dialog"):
    @st.dialog("🏛️ Detalhes do Ambiente de Inovação")
    def exibir_popup_ambiente(dados):
        st.subheader(dados["nome"])
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**ID SPAI:** `{dados['id']}`")
            st.markdown(f"**Categoria:** {dados['categoria']}")
            st.markdown(f"**Tipo:** {dados['tipo']}")
        with c2:
            st.markdown(f"**Município:** {dados['municipio']}")
            st.markdown(f"**Região Administrativa:** {dados['regiao']}")
            
        st.markdown("---")
        
        if dados["verticais"]:
            st.markdown("**💻 Verticais Tecnológicas Mapeadas:**")
            st.write(", ".join(dados["verticais"]))
        else:
            st.caption("ℹ️ Verticais Tecnológicas: Não mapeadas na fase atual.")

        if dados["segmentos"]:
            st.markdown("**🏢 Segmentos de Mercado Mapeados:**")
            st.write(", ".join(dados["segmentos"]))
        else:
            st.caption("ℹ️ Segmentos de Mercado: Não mapeados na fase atual.")

# Mapeamento mestre de dados do SPAI pelo Nome do Ambiente
col_spai_id = df_spai.columns[0]
col_spai_cat = df_spai.columns[1] if df_spai.shape[1] > 1 else None
col_spai_reg = df_spai.columns[2] if df_spai.shape[1] > 2 else None
col_spai_mun = df_spai.columns[3] if df_spai.shape[1] > 3 else None
col_spai_tipo = df_spai.columns[4] if df_spai.shape[1] > 4 else None
col_spai_nome = df_spai.columns[5] if df_spai.shape[1] > 5 else df_spai.columns[0]

MAPA_SPAI_DADOS = {}
for _, r in df_spai.iterrows():
    key = normalizar_texto(r[col_spai_nome])
    val_id = pd.to_numeric(r[col_spai_id], errors="coerce")
    MAPA_SPAI_DADOS[key] = {
        "ID": int(val_id) if pd.notna(val_id) else 0,
        "Categoria": str(r[col_spai_cat]).strip().title() if col_spai_cat and pd.notna(r[col_spai_cat]) else "",
        "Regiao": str(r[col_spai_reg]).strip() if col_spai_reg and pd.notna(r[col_spai_reg]) else "",
        "Municipio": str(r[col_spai_mun]).strip() if col_spai_mun and pd.notna(r[col_spai_mun]) else "",
        "Tipo": str(r[col_spai_tipo]).strip().title() if col_spai_tipo and pd.notna(r[col_spai_tipo]) else ""
    }

# ----------------------------------------------------
# BARRA LATERAL - NAVEGAÇÃO ENTRE PAINÉIS
# ----------------------------------------------------

# Logo do CTI exibido no TOPO da barra lateral (ACIMA do logo da SCTI)
if os.path.exists(CAMINHO_CTI):
    st.sidebar.image(CAMINHO_CTI, width='stretch')

# Logo da SCTI exibido ACIMA do logo do SPAI
if os.path.exists(CAMINHO_LOGO_SCTI):
    st.sidebar.image(CAMINHO_LOGO_SCTI, width='stretch')

# Logo Horizontal do SPAI
if os.path.exists(CAMINHO_LOGO_HORIZONTAL):
    st.sidebar.image(CAMINHO_LOGO_HORIZONTAL, width='stretch')

painel_selecionado = st.sidebar.radio(
    "Navegação por Painel:",
    ["Visão Geral", "Verticais Tecnológicas", "Segmentos de Mercado"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.info("""
**Sistema Paulista de Ambientes de Inovação (SPAI)**

O **SPAI** é um programa da política pública de inovação do Estado de São Paulo, gerido pela **Secretaria de Ciência, Tecnologia e Inovação (SCTI-SP)**.

📌 **Objetivo:**
Incentivar investimentos em inovação, fortalecer a pesquisa científica e conectar parques tecnológicos, incubadoras, centros de inovação, centros de conexão e aceleradoras ao setor produtivo.

🔗 [Saiba mais no site oficial da SCTI](https://www.inovacao.sp.gov.br/sec_tecnologia_inovacao/programas/spai)
""")

# Logo do Projeto Japi exibido logo abaixo da caixa informativa
if os.path.exists(CAMINHO_LOGO_JAPI):
    st.sidebar.image(CAMINHO_LOGO_JAPI, width='stretch')

# Função genérica de renderização de painel fato com interatividade Treemap -> Tabela
def renderizar_painel_fato(df_fato, titulo_item, item_plural, explicacao_conceitual):
    st.title(f"{item_plural}")
    st.caption(f"Mapeamento e distribuição de {item_plural.lower()} nos ambientes de inovação de São Paulo")

    # Bloco Explicativo do Significado Prático
    st.markdown(f"""
    <div class="concept-card">
        <div class="concept-title">💡 Significado Prático para Inovação e Empreendedorismo</div>
        <div class="concept-text">{explicacao_conceitual}</div>
    </div>
    """, unsafe_allow_html=True)

    st.info(f"ℹ️ **Mapeamento 2025:** Os dados referentes a **{item_plural}** contemplam os **68 ambientes de inovação** participantes do Mapeamento 2025. Os dados dos demais ambientes serão incorporados após a conclusão do Mapeamento 2026.")

    st.markdown("---")

    if df_fato is None or df_fato.empty:
        st.warning(f"⚠️ A planilha de {item_plural} não foi localizada no caminho especificado.")
        uploaded_file = st.file_uploader(f"Enviar arquivo de {item_plural}", type=["xlsx", "xls"], key=f"up_{titulo_item}")
        if uploaded_file is not None:
            try:
                df_fato = pd.read_excel(uploaded_file)
                st.success(f"Arquivo de {item_plural} carregado com sucesso!")
            except Exception as e:
                st.error(f"Erro ao processar o arquivo enviado: {e}")

    if df_fato is not None and not df_fato.empty:
        cols = list(df_fato.columns)

        def encontrar_coluna(palavras):
            for p in palavras:
                for col in cols:
                    if p.lower() in str(col).lower():
                        return col
            return None

        col_id_fato = encontrar_coluna(["id", "código", "codigo"]) or cols[0]
        col_nome = encontrar_coluna(["nome", "ambiente", "instituição", "instituicao"]) or (cols[1] if len(cols) > 1 else cols[0])
        col_atuacao = encontrar_coluna(["vertical", "segmento", "área", "area", "atuação", "atuacao"]) or cols[-1]

        # Enriquecimento com os dados cruzados da Base Geral SPAI
        df_base = df_fato.copy()

        def extrair_dado_spai(row, propriedade):
            key = normalizar_texto(row[col_nome])
            if key in MAPA_SPAI_DADOS and MAPA_SPAI_DADOS[key].get(propriedade):
                return MAPA_SPAI_DADOS[key][propriedade]
            return ""

        df_base["SPAI_ID"] = df_base.apply(lambda r: extrair_dado_spai(r, "ID"), axis=1)
        df_base["SPAI_Categoria"] = df_base.apply(lambda r: extrair_dado_spai(r, "Categoria"), axis=1)
        df_base["SPAI_Regiao"] = df_base.apply(lambda r: extrair_dado_spai(r, "Regiao"), axis=1)
        df_base["SPAI_Municipio"] = df_base.apply(lambda r: extrair_dado_spai(r, "Municipio"), axis=1)
        df_base["SPAI_Tipo"] = df_base.apply(lambda r: extrair_dado_spai(r, "Tipo"), axis=1)

        # ----------------------------------------------------
        # FILTROS CENTRALIZADOS NA PÁGINA
        # ----------------------------------------------------
        st.markdown("<h4 style='text-align: center; color: #005f73; margin-bottom: 12px;'>🔍 Filtros de Ambientes</h4>", unsafe_allow_html=True)

        cats_unicas = [c for c in df_base["SPAI_Categoria"].unique() if c != ""]
        cats_ordenadas = ordenar_categorias_prioritarias(cats_unicas)
        cats_opts = ["Todas"] + cats_ordenadas

        col_f1, col_f2, col_f3 = st.columns([1, 8, 1])
        with col_f2:
            st.markdown("<div style='text-align: center; font-size: 13px; font-weight: 600; color: #334155; margin-bottom: 4px;'>Categoria do Ambiente:</div>", unsafe_allow_html=True)
            if hasattr(st, "pills"):
                cat_sel = st.pills(f"cat_{titulo_item}", cats_opts, default="Todas", label_visibility="collapsed")
            else:
                cat_sel = st.selectbox(f"Categoria ({titulo_item}):", cats_opts, label_visibility="collapsed")

        df_filt_f = df_base.copy()
        if cat_sel and cat_sel != "Todas":
            df_filt_f = df_filt_f[df_filt_f["SPAI_Categoria"] == cat_sel]

        # Dropdowns Secundários
        f_cols = st.columns(4)

        atuacao_list = ["Todos"] + sorted([a for a in df_filt_f[col_atuacao].dropna().astype(str).str.strip().unique() if a != ""])
        item_sel = f_cols[0].selectbox(f"{titulo_item}:", atuacao_list, key=f"item_{titulo_item}")
        if item_sel != "Todos":
            df_filt_f = df_filt_f[df_filt_f[col_atuacao].astype(str).str.strip() == item_sel]

        reg_list = ["Todas"] + sorted([r for r in df_filt_f["SPAI_Regiao"].unique() if r != ""])
        reg_sel = f_cols[1].selectbox("Região Administrativa:", reg_list, key=f"reg_{titulo_item}")
        if reg_sel != "Todas":
            df_filt_f = df_filt_f[df_filt_f["SPAI_Regiao"] == reg_sel]

        mun_list = ["Todos"] + sorted([m for m in df_filt_f["SPAI_Municipio"].unique() if m != ""])
        mun_sel = f_cols[2].selectbox("Município:", mun_list, key=f"mun_{titulo_item}")
        if mun_sel != "Todos":
            df_filt_f = df_filt_f[df_filt_f["SPAI_Municipio"] == mun_sel]

        tipo_list = ["Todos"] + sorted([t for t in df_filt_f["SPAI_Tipo"].unique() if t != ""])
        tipo_sel = f_cols[3].selectbox("Tipo (Pleno / Associado):", tipo_list, key=f"tipo_{titulo_item}")
        if tipo_sel != "Todos":
            df_filt_f = df_filt_f[df_filt_f["SPAI_Tipo"] == tipo_sel]

        st.markdown("---")

        # ----------------------------------------------------
        # LINHA SUPERIOR DE CARDS KPI
        # ----------------------------------------------------
        total_amb = df_filt_f[col_nome].nunique()
        total_items = df_filt_f[col_atuacao].nunique()
        total_muns = df_filt_f["SPAI_Municipio"].replace("", None).dropna().nunique()
        if total_muns == 0:
            total_muns = total_amb
        media_items = round(len(df_filt_f) / max(total_amb, 1), 1)

        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(f'<div class="metric-card"><div class="metric-val">{total_amb}</div><div class="metric-lbl">Ambientes Mapeados</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="metric-card"><div class="metric-val">{total_items}</div><div class="metric-lbl">{item_plural}</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="metric-card"><div class="metric-val">{total_muns}</div><div class="metric-lbl">Municípios Atendidos</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="metric-card"><div class="metric-val">{media_items}</div><div class="metric-lbl">Média / Ambiente</div></div>', unsafe_allow_html=True)

        st.markdown("---")

        # ----------------------------------------------------
        # 1º: TREEMAP (DISTRIBUIÇÃO DOS AMBIENTES)
        # ----------------------------------------------------
        st.subheader(f"📊 Distribuição dos Ambientes por {titulo_item}")
        st.caption("💡 *Clique em qualquer bloco do gráfico para filtrar a tabela detalhada abaixo:*")

        df_dist = df_filt_f.groupby(col_atuacao)[col_nome].nunique().reset_index()
        df_dist.columns = [titulo_item, "Quantidade de Ambientes"]
        df_dist = df_dist.sort_values(by="Quantidade de Ambientes", ascending=False)

        item_clicado_grafico = None

        if not df_dist.empty:
            escala_neutra = ["#94d2bd", "#0a9396", "#005f73", "#1e293b"]

            fig_tree = px.treemap(
                df_dist,
                path=[titulo_item],
                values="Quantidade de Ambientes",
                color="Quantidade de Ambientes",
                color_continuous_scale=escala_neutra
            )

            fig_tree.update_traces(
                textinfo="label+value",
                textfont=dict(size=13, family="Inter, sans-serif"),
                hovertemplate="<b>%{label}</b><br>Ambientes Atendidos: <b>%{value}</b><extra></extra>",
                marker=dict(cornerradius=4)
            )

            fig_tree.update_layout(
                coloraxis_showscale=False,
                height=480,
                margin={"l": 10, "r": 10, "t": 20, "b": 20},
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Inter, sans-serif", color="#334155")
            )

            # Evento de clique habilitado no Treemap
            tree_event = st.plotly_chart(
                fig_tree,
                on_select="rerun",
                selection_mode="points",
                key=f"tree_chart_{titulo_item}",
                width='stretch'
            )

            # Extração do bloco clicado
            if tree_event and isinstance(tree_event, dict) and "selection" in tree_event:
                pts = tree_event["selection"].get("points", [])
                if pts:
                    p0 = pts[0]
                    lbl = p0.get("label") or p0.get("id")
                    if lbl:
                        item_clicado_grafico = str(lbl).strip()
                    else:
                        idx_p = p0.get("point_index")
                        if idx_p is not None and idx_p < len(df_dist):
                            item_clicado_grafico = str(df_dist.iloc[idx_p][titulo_item]).strip()

        st.markdown("---")

        # ----------------------------------------------------
        # 2º: TABELA DETALHADA INTERATIVA
        # ----------------------------------------------------
        df_table_source = df_filt_f.copy()

        if item_clicado_grafico:
            st.subheader(f"📋 Relação Completa de Ambientes em **{item_clicado_grafico}**")
            st.info(f"🎯 **Filtrado pelo Treemap:** Exibindo os ambientes que atuam em **{item_clicado_grafico}**. *(Clique no bloco do gráfico novamente para limpar)*")
            df_table_source = df_table_source[df_table_source[col_atuacao].astype(str).str.strip() == item_clicado_grafico]
        else:
            st.subheader(f"📋 Relação Completa de Ambientes e {item_plural} Atendidos")

        df_grouped = df_table_source.groupby(["SPAI_ID", col_nome], dropna=False).agg(
            Qtd_Itens=(col_atuacao, "nunique"),
            Itens_Atendidos=(col_atuacao, lambda x: ", ".join(sorted(x.dropna().astype(str).str.strip().unique())))
        ).reset_index()

        col_qtd_nome = f"Nº de {item_plural}"
        col_renames = {
            "SPAI_ID": "ID",
            col_nome: "Ambiente de Inovação",
            "Qtd_Itens": col_qtd_nome,
            "Itens_Atendidos": f"{item_plural} Atendidos"
        }

        df_grouped = df_grouped.rename(columns=col_renames)
        df_grouped["ID"] = pd.to_numeric(df_grouped["ID"], errors="coerce").fillna(0).astype(int)
        df_grouped = df_grouped.set_index("ID")

        # Ordenação Decrescente pelo número de verticais / segmentos
        df_grouped = df_grouped.sort_values(by=col_qtd_nome, ascending=False)

        configs = {
            "Ambiente de Inovação": st.column_config.TextColumn("Ambiente de Inovação", width="medium"),
            col_qtd_nome: st.column_config.NumberColumn(col_qtd_nome, format="%d", width="small"),
            f"{item_plural} Atendidos": st.column_config.TextColumn(f"{item_plural} Atendidos", width="large")
        }

        st.dataframe(
            df_grouped,
            column_config=configs,
            width='stretch',
            height=380
        )


# ====================================================
# PAINEL 1: VISÃO GERAL
# ====================================================
if painel_selecionado == "Visão Geral":
    st.title("Visão Geral dos Ambientes de Inovação")
    st.caption("Sistema Paulista de Ambientes de Inovação — Secretaria de Ciência, Tecnologia e Inovação de SP")
    st.write("")

    total_membros = len(df_spai)

    if df_spai.shape[1] > 4:
        tipo_col = df_spai.iloc[:, 4].astype(str).str.strip().str.title()
        plenos = (tipo_col == 'Pleno').sum()
        associados = (tipo_col == 'Associado').sum()
    else:
        plenos, associados = 0, 0

    regioes_count = df_spai.iloc[:, 2].dropna().astype(str).str.strip().nunique() if df_spai.shape[1] > 2 else 0

    if df_spai.shape[1] > 1:
        cat_col = df_spai.iloc[:, 1].astype(str).str.strip().str.upper()
        parques = cat_col.str.contains('PARQUE', na=False).sum()
        incubadoras = cat_col.str.contains('INCUBADORA', na=False).sum()
        centros = cat_col.str.contains('CENTRO DE INOV', na=False).sum()
        centros_conexao = (cat_col.str.contains('HUB', na=False) | cat_col.str.contains('CONEX', na=False)).sum()
        aceleradoras = cat_col.str.contains('ACELERADORA', na=False).sum()
        outros = total_membros - (parques + incubadoras + centros + centros_conexao + aceleradoras)
        if outros < 0: outros = 0
    else:
        parques = incubadoras = centros = centros_conexao = aceleradoras = outros = 0

    top1, top2, top3, top4 = st.columns(4)
    top1.markdown(f'''
        <div class="kpi-box-membros">
            <div class="kpi-val-membros">{total_membros}</div>
            <div class="kpi-lbl-membros">Total de Membros SPAI</div>
        </div>
    ''', unsafe_allow_html=True)
    top2.markdown(f'<div class="kpi-box-top"><div class="kpi-val-top">{plenos}</div><div class="kpi-lbl-top">Membros Plenos</div></div>', unsafe_allow_html=True)
    top3.markdown(f'<div class="kpi-box-top"><div class="kpi-val-top">{associados}</div><div class="kpi-lbl-top">Membros Associados</div></div>', unsafe_allow_html=True)
    top4.markdown(f'<div class="kpi-box-top"><div class="kpi-val-top">{regioes_count}</div><div class="kpi-lbl-top">Regiões Administrativas</div></div>', unsafe_allow_html=True)

    st.write("")

    sec1, sec2, sec3, sec4, sec5, sec6, sec7 = st.columns(7)
    sec1.markdown(f'<div class="kpi-box-sec"><div class="kpi-val-sec">{parques}</div><div class="kpi-lbl-sec">Parques Tecnológicos</div></div>', unsafe_allow_html=True)
    sec2.markdown(f'<div class="kpi-box-sec"><div class="kpi-val-sec">{incubadoras}</div><div class="kpi-lbl-sec">Incubadoras</div></div>', unsafe_allow_html=True)
    sec3.markdown(f'<div class="kpi-box-sec"><div class="kpi-val-sec">{centros}</div><div class="kpi-lbl-sec">Centros de Inovação</div></div>', unsafe_allow_html=True)
    sec4.markdown(f'<div class="kpi-box-sec"><div class="kpi-val-sec">{centros_conexao}</div><div class="kpi-lbl-sec">Centros de Conexão</div></div>', unsafe_allow_html=True)
    sec5.markdown(f'<div class="kpi-box-sec"><div class="kpi-val-sec">{aceleradoras}</div><div class="kpi-lbl-sec">Aceleradoras</div></div>', unsafe_allow_html=True)
    sec6.markdown(f'<div class="kpi-box-sec"><div class="kpi-val-sec">{outros}</div><div class="kpi-lbl-sec">Outros Formatos</div></div>', unsafe_allow_html=True)
    sec7.markdown('<div class="kpi-box-sec"><div class="kpi-lbl-sec">Distritos de Inovação</div><div class="kpi-sub-sec">em formação</div></div>', unsafe_allow_html=True)

    st.markdown("---")

    # ----------------------------------------------------
    # FILTROS E SEGMENTADORES DE DADOS (COM DESTAQUE VISUAL EXECUTIVO)
    # ----------------------------------------------------
    st.subheader("🔍 Filtros e Segmentadores")

    st.markdown(f"""
    <div class="filter-card-tipo">
        <div class="filter-card-hdr">🏛️ PASSO 1: Selecione o Tipo de Membro SPAI</div>
        <div class="filter-card-sub">Escolha a categoria principal de vinculação para atualizar dinamicamente as opções de formato de ambiente abaixo:</div>
    </div>
    """, unsafe_allow_html=True)

    rotulo_todos = f"🌐 Todos os Membros ({total_membros})"
    rotulo_plenos = f"🏛️ Membros Plenos ({plenos})"
    rotulo_associados = f"🤝 Membros Associados ({associados})"
    opcoes_tipo_display = [rotulo_todos, rotulo_plenos, rotulo_associados]

    if hasattr(st, "pills"):
        tipo_sel_display = st.pills("tipo_membro_highlight", opcoes_tipo_display, default=rotulo_todos, label_visibility="collapsed")
    else:
        tipo_sel_display = st.radio("Tipo de Membro:", opcoes_tipo_display, horizontal=True)

    # Filtragem primária pelo Tipo de Membro
    df_filt = df_spai.copy()
    if "Plenos" in str(tipo_sel_display) and df_filt.shape[1] > 4:
        df_filt = df_filt[df_filt.iloc[:, 4].astype(str).str.strip().str.title() == "Pleno"]
    elif "Associados" in str(tipo_sel_display) and df_filt.shape[1] > 4:
        df_filt = df_filt[df_filt.iloc[:, 4].astype(str).str.strip().str.title() == "Associado"]

    # Extração dinâmica de categorias disponíveis resultantes do Tipo de Membro
    col_b_categorias = df_filt.iloc[:, 1].dropna().astype(str).str.strip().unique().tolist()
    cats_formatadas = list(set([c.title() for c in col_b_categorias if c != ""]))
    cats_ordenadas = ordenar_categorias_prioritarias(cats_formatadas)
    opcoes_categorias = ["Todas"] + cats_ordenadas

    st.write("")
    st.markdown("**📂 PASSO 2: Categoria / Formato do Ambiente:**")
    if hasattr(st, "pills"):
        cat_sel = st.pills(
            label="Categoria do Ambiente",
            options=opcoes_categorias,
            default="Todas",
            key="cat_pills_vg",
            label_visibility="collapsed"
        )
    else:
        cat_sel = st.selectbox("Categoria:", opcoes_categorias, key="cat_select_vg")

    if cat_sel and cat_sel != "Todas":
        df_filt = df_filt[df_filt.iloc[:, 1].astype(str).str.strip().str.title() == cat_sel]

    st.write("")

    st.markdown("**📍 PASSO 3: Filtros de Localização Regional:**")
    c1, c2 = st.columns(2)

    regioes_unicas = ["Todas"] + sorted(list(df_filt.iloc[:, 2].dropna().astype(str).str.strip().unique())) if df_filt.shape[1] > 2 else ["Todas"]
    regiao_sel = c1.selectbox("Região Administrativa", regioes_unicas)
    if regiao_sel != "Todas" and df_filt.shape[1] > 2:
        df_filt = df_filt[df_filt.iloc[:, 2].astype(str).str.strip() == regiao_sel]

    cidades_unicas = ["Todas"] + sorted(list(df_filt.iloc[:, 3].dropna().astype(str).str.strip().unique())) if df_filt.shape[1] > 3 else ["Todas"]
    cidade_sel = c2.selectbox("Município", cidades_unicas)
    if cidade_sel != "Todas" and df_filt.shape[1] > 3:
        df_filt = df_filt[df_filt.iloc[:, 3].astype(str).str.strip() == cidade_sel]

    st.write("")
    col_m, col_l = st.columns([6, 4])

    col_id_orig = df_filt.columns[0]
    col_nome_orig = df_filt.columns[5] if df_filt.shape[1] > 5 else df_filt.columns[0]

    with col_m:
        # Tratamento para coordenadas no mapa
        df_m = df_filt.dropna(subset=["Lat", "Lon"]).copy()
        df_m["Lat"] = pd.to_numeric(df_m["Lat"], errors="coerce")
        df_m["Lon"] = pd.to_numeric(df_m["Lon"], errors="coerce")
        df_m = df_m.dropna(subset=["Lat", "Lon"])

        if not df_m.empty:
            hover_dict = {"Lat": False, "Lon": False}
            labels_dict = {}
            if df_filt.shape[1] > 0: hover_dict[df_filt.columns[0]] = True; labels_dict[df_filt.columns[0]] = "ID"
            if df_filt.shape[1] > 1: hover_dict[df_filt.columns[1]] = True; labels_dict[df_filt.columns[1]] = "Categoria"
            if df_filt.shape[1] > 2: hover_dict[df_filt.columns[2]] = True; labels_dict[df_filt.columns[2]] = "Região Administrativa"
            if df_filt.shape[1] > 3: hover_dict[df_filt.columns[3]] = True; labels_dict[df_filt.columns[3]] = "Município"
            if df_filt.shape[1] > 4: hover_dict[df_filt.columns[4]] = True; labels_dict[df_filt.columns[4]] = "Tipo"

            fig = px.scatter_map(
                df_m, lat="Lat", lon="Lon",
                hover_name=col_nome_orig,
                hover_data=hover_dict,
                labels=labels_dict,
                zoom=6.1, center={"lat": -22.3, "lon": -48.8},
                height=480
            )
            fig.update_layout(
                map_style="carto-positron",
                margin={"r":0, "t":0, "l":0, "b":0}
            )
            fig.update_traces(marker=dict(size=14, color="#005f73", opacity=0.85))

            # Captura eventos de seleção ao clicar nos pontos do mapa
            map_event = st.plotly_chart(
                fig,
                on_select="rerun",
                selection_mode="points",
                width='stretch'
            )

            # Modal popup de detalhes ao clicar no ponto do mapa
            if map_event and isinstance(map_event, dict) and "selection" in map_event:
                pts = map_event["selection"].get("points", [])
                if pts:
                    idx_ponto = pts[0].get("point_index")
                    if idx_ponto is not None and idx_ponto < len(df_m):
                        row_sel = df_m.iloc[idx_ponto]
                        
                        nome_env = str(row_sel[col_nome_orig]).strip()
                        id_env = int(pd.to_numeric(row_sel[col_id_orig], errors="coerce") or 0)
                        cat_env = str(row_sel.iloc[1]).strip().title() if df_filt.shape[1] > 1 else ""
                        reg_env = str(row_sel.iloc[2]).strip() if df_filt.shape[1] > 2 else ""
                        mun_env = str(row_sel.iloc[3]).strip() if df_filt.shape[1] > 3 else ""
                        tipo_env = str(row_sel.iloc[4]).strip().title() if df_filt.shape[1] > 4 else ""
                        
                        v_env, s_env = buscar_detalhes_ambiente(nome_env)
                        
                        dados_popup = {
                            "nome": nome_env,
                            "id": id_env,
                            "categoria": cat_env,
                            "regiao": reg_env,
                            "municipio": mun_env,
                            "tipo": tipo_env,
                            "verticais": v_env,
                            "segmentos": s_env
                        }
                        
                        if hasattr(st, "dialog"):
                            exibir_popup_ambiente(dados_popup)
                        else:
                            with st.expander(f"📍 Detalhes de {nome_env}", expanded=True):
                                st.write(f"**ID SPAI:** `{id_env}` | **Categoria:** {cat_env} | **Tipo:** {tipo_env}")
                                st.write(f"**Município:** {mun_env} | **Região:** {reg_env}")
                                if v_env: st.write("**Verticais:** " + ", ".join(v_env))
                                if s_env: st.write("**Segmentos:** " + ", ".join(s_env))
        else:
            st.warning("⚠️ Não há coordenadas geográficas válidas para exibir no mapa com os filtros selecionados.")

    with col_l:
        st.markdown("<h4 style='text-align: center;'>Nome do Ambiente</h4>", unsafe_allow_html=True)

        df_exibicao = df_filt[[col_id_orig, col_nome_orig]].copy()
        df_exibicao.columns = ["ID", "Nome do Ambiente"]
        df_exibicao["ID"] = pd.to_numeric(df_exibicao["ID"], errors="coerce").fillna(0).astype(int)
        df_exibicao = df_exibicao.set_index("ID")

        st.dataframe(
            df_exibicao,
            column_config={
                "Nome do Ambiente": st.column_config.TextColumn("Nome do Ambiente", width="large")
            },
            width='stretch',
            height=430
        )

# ====================================================
# PAINEL 2: VERTICAIS TECNOLÓGICAS
# ====================================================
elif painel_selecionado == "Verticais Tecnológicas":
    explicacao_verticais = (
        "As <b>Verticais Tecnológicas</b> representam os <b>domínios de conhecimento técnico, ferramentas e tecnologias habilitadoras</b> "
        "(como <i>Inteligência Artificial, Biotecnologia, IoT, Robótica, Nanotecnologia, Big Data</i>) desenvolvidos ou aplicados nos ambientes de inovação.<br><br>"
        "🎯 <b>Impacto no Empreendedorismo:</b> Indicam a <b>capacidade tecnológica e científica</b> que os ecossistemas locais possuem para apoiar startups, gerar patentes, "
        "promover a transferência de tecnologia da academia para a indústria e solucionar desafios complexos por meio de inovação profunda (<i>Deep Techs</i>)."
    )
    renderizar_painel_fato(df_verticais, "Vertical Tecnológica", "Verticais Tecnológicas", explicacao_verticais)

# ====================================================
# PAINEL 3: SEGMENTOS DE MERCADO
# ====================================================
elif painel_selecionado == "Segmentos de Mercado":
    explicacao_segmentos = (
        "Os <b>Segmentos de Mercado</b> correspondem aos <b>setores econômicos, indústrias e cadeias de valor finais</b> "
        "(como <i>Agronegócio, Saúde, Educação, Finanças, Logística, Energia, Cidades Inteligentes</i>) atendidos pelas soluções do ecossistema.<br><br>"
        "🎯 <b>Impacto no Empreendedorismo:</b> Definem o <b>mercado consumidor, a demanda real e o modelo de negócios</b> das empresas e startups residentes. "
        "Ajudam a identificar oportunidades de <i>Open Innovation</i> (Inovação Aberta) com grandes corporações e a direcionar novos negócios para mercados com alto potencial de tração e escala."
    )
    renderizar_painel_fato(df_segmentos, "Segmento de Mercado", "Segmentos de Mercado", explicacao_segmentos)

# ----------------------------------------------------
# RODAPÉ
# ----------------------------------------------------
st.markdown("---")
f1, f2 = st.columns([1, 1])
f1.caption("Fontes: Mapeamento do Sistema Paulista de Ambientes de Inovação (SPAI).")
f2.caption(f"Última verificação: {datetime.now(FUSO_BRASILIA).strftime('%d/%m/%Y %H:%M')}")
