import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import unicodedata
from datetime import datetime, timezone, timedelta

# Caminhos exatos no GitHub / Repositório
CAMINHO_SEGMENTOS = "data/Fato_Segmentos.xlsx"
CAMINHO_VERTICAIS = "data/Fato_Verticais.xlsx"
CAMINHO_CTI = "assets/CTI.png"
CAMINHO_LOGO_HORIZONTAL = "assets/Logotipo Horizontal SPAI.png"
CAMINHO_LOGO_SPAI_VERTICAL = "assets/SPAI_Vertical.png"
CAMINHO_LOGO_JAPI = "assets/Logo_JAPI.png"

# Fuso horário oficial de Brasília (UTC-3)
FUSO_BRASILIA = timezone(timedelta(hours=-3))

# Paleta de Cores Institucionais SPAI / SCTI-SP
PALETA_SPAI_ESCALA = ["#94d2bd", "#0a9396", "#005f73", "#1e293b"]

# Função para normalizar strings (remove acentos, maiúsculas e espaços extras)
def normalizar_texto(txt):
    if not isinstance(txt, str):
        txt = str(txt) if pd.notna(txt) else ""
    txt = txt.strip().lower()
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('utf-8')
    return txt

# Retorna cor institucional sóbria e forma geométrica por categoria de ambiente
def obter_estilo_categoria_mapa(categoria):
    c_norm = normalizar_texto(categoria)
    if "parque" in c_norm:
        return {"cor": "#1B365D", "simbolo": "●", "forma": "circle"}
    if "incubadora" in c_norm:
        return {"cor": "#2E7D7B", "simbolo": "◆", "forma": "diamond"}
    if "centro de inov" in c_norm:
        return {"cor": "#C86D1F", "simbolo": "■", "forma": "square"}
    if "centro de conex" in c_norm or "hub" in c_norm:
        return {"cor": "#4A6FA5", "simbolo": "▲", "forma": "triangle"}
    return {"cor": "#64748B", "simbolo": "●", "forma": "dot"}

# Gera o ícone SVG geométrico da legenda
def gerar_svg_forma_legenda(forma, cor):
    if forma == "circle":
        return f'<svg width="11" height="11" viewBox="0 0 12 12" style="flex-shrink:0;"><circle cx="6" cy="6" r="5" fill="{cor}"/></svg>'
    if forma == "diamond":
        return f'<svg width="11" height="11" viewBox="0 0 12 12" style="flex-shrink:0;"><polygon points="6,1 11,6 6,11 1,6" fill="{cor}"/></svg>'
    if forma == "square":
        return f'<svg width="11" height="11" viewBox="0 0 12 12" style="flex-shrink:0;"><rect x="1.5" y="1.5" width="9" height="9" rx="1.5" fill="{cor}"/></svg>'
    if forma == "triangle":
        return f'<svg width="11" height="11" viewBox="0 0 12 12" style="flex-shrink:0;"><polygon points="6,1.5 11,10.5 1,10.5" fill="{cor}"/></svg>'
    return f'<svg width="10" height="10" viewBox="0 0 12 12" style="flex-shrink:0;"><circle cx="6" cy="6" r="4.5" fill="{cor}"/></svg>'

# Formata percentual no padrão brasileiro (ex: 31,5% ou 12%)
def formatar_percentual_ptbr(valor):
    arredondado = round(float(valor), 1)
    if arredondado == int(arredondado):
        return f"{int(arredondado)}%"
    return f"{arredondado:.1f}%".replace(".", ",")

# Retorna ícone SVG minimalista de acordo com o tema da Vertical ou Segmento
def obter_icone_svg(categoria, cor="#005f73", tamanho=26):
    cat_norm = normalizar_texto(categoria)
    props = f'width="{tamanho}" height="{tamanho}" viewBox="0 0 24 24" fill="none" stroke="{cor}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"'

    if cat_norm == "outros":
        return f'<svg {props}><circle cx="12" cy="12" r="1.5" fill="{cor}"/><circle cx="19" cy="12" r="1.5" fill="{cor}"/><circle cx="5" cy="12" r="1.5" fill="{cor}"/></svg>'
    if any(k in cat_norm for k in ["saude", "vida", "bem-estar", "bem estar", "medic", "farma", "hospital"]):
        return f'<svg {props}><path d="M19 14c1.49-1.46 3-3.21 3-5.5A5.5 5.5 0 0 0 16.5 3c-1.76 0-3 .5-4.5 2-1.5-1.5-2.74-2-4.5-2A5.5 5.5 0 0 0 2 8.5c0 2.3 1.5 4.05 3 5.5l7 7Z"/><path d="M3.22 12H9.5l.5-1 2 4.5 2-7 1.5 3.5h5.27"/></svg>'
    if any(k in cat_norm for k in ["agro", "aliment", "bioeconomia", "rural", "agric", "florest"]):
        return f'<svg {props}><path d="M7 20h10"/><path d="M10 20c5.5-2.5.8-6.4 3-10"/><path d="M9.5 9.4c1.1.8 1.8 2.2 2.3 3.7-2 .4-3.5.4-4.8-.3-1.2-.6-2.3-1.9-3-4.2 2.8-.5 4.4 0 5.5.8z"/><path d="M14.1 6a7 7 0 0 0-1.1 4c1.9-.1 3.3-.6 4.3-1.4 1-1 1.6-2.3 1.7-4.6-2.7.1-4 1-4.9 2z"/></svg>'
    if any(k in cat_norm for k in ["industr", "manufatura", "materiais", "quimic", "metal", "mecan"]):
        return f'<svg {props}><path d="M2 20a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V8l-7 5V8l-7 5V4a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2Z"/><path d="M17 18h1"/><path d="M12 18h1"/><path d="M7 18h1"/></svg>'
    if any(k in cat_norm for k in ["energia", "clima", "sustentab", "ambient", "solar", "limpa", "verde"]):
        return f'<svg {props}><circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/></svg>'
    if any(k in cat_norm for k in ["educa", "ciencia", "publico", "ensino", "governo", "social"]):
        return f'<svg {props}><path d="M21.42 10.922a1 1 0 0 0-.019-1.838L12.83 5.18a2 2 0 0 0-1.66 0L2.6 9.08a1 1 0 0 0 0 1.832l8.57 3.908a2 2 0 0 0 1.66 0z"/><path d="M22 10v6"/><path d="M6 12.5V16a6 3 0 0 0 12 0v-3.5"/></svg>'
    if any(k in cat_norm for k in ["inteligencia artificial", "ia ", " ia", "dados", "big data", "computa", "algoritm"]):
        return f'<svg {props}><rect x="4" y="4" width="16" height="16" rx="2"/><rect x="9" y="9" width="6" height="6"/><path d="M15 2v2"/><path d="M15 20v2"/><path d="M2 15h2"/><path d="M2 9h2"/><path d="M20 15h2"/><path d="M20 9h2"/><path d="M9 2v2"/><path d="M9 20v2"/></svg>'
    if any(k in cat_norm for k in ["tecnologia", "digital", "software", "tic", "sistema", "cloud", "nuvem", "web"]):
        return f'<svg {props}><path d="M20 16V7a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v9m16 0H4m16 0 1.28 2.55a1 1 0 0 1-.9 1.45H3.62a1 1 0 0 1-.9-1.45L4 16"/></svg>'
    if any(k in cat_norm for k in ["infraestrutura", "logistica", "mobilidade", "transport", "cidade", "urbana"]):
        return f'<svg {props}><path d="M14 18V6a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v11a1 1 0 0 0 1 1h2"/><path d="M15 18H9"/><path d="M19 18h2a1 1 0 0 0 1-1v-3.65a1 1 0 0 0-.22-.624l-3.48-4.35A1 1 0 0 0 17.52 8H14"/><circle cx="17" cy="18" r="2"/><circle cx="7" cy="18" r="2"/></svg>'
    if any(k in cat_norm for k in ["defesa", "seguranca", "espaco", "aero", "ciber"]):
        return f'<svg {props}><path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/><path d="m9 12 2 2 4-4"/></svg>'
    if any(k in cat_norm for k in ["financ", "seguro", "econom", "fintech", "banco", "credito"]):
        return f'<svg {props}><line x1="3" x2="21" y1="22" y2="22"/><line x1="6" x2="6" y1="18" y2="11"/><line x1="10" x2="10" y1="18" y2="11"/><line x1="14" x2="14" y1="18" y2="11"/><line x1="18" x2="18" y1="18" y2="11"/><polygon points="12 2 20 7 4 7"/></svg>'
    if any(k in cat_norm for k in ["biotecnologia", "genet", "biolog", "nano", "atom"]):
        return f'<svg {props}><circle cx="12" cy="12" r="1"/><path d="M20.2 20.2c2.04-2.03.02-7.36-4.5-11.9-4.54-4.52-9.87-6.54-11.9-4.5-2.04 2.03-.02 7.36 4.5 11.9 4.54 4.52 9.87 6.54 11.9 4.5Z"/><path d="M15.7 15.7c4.52-4.54 6.54-9.87 4.5-11.9-2.03-2.04-7.36-.02-11.9 4.5-4.52 4.54-6.54 9.87-4.5 11.9 2.03 2.04 7.36.02 11.9-4.5Z"/></svg>'
    if any(k in cat_norm for k in ["iot", "internet das coisas", "conectiv", "telecom", "sensor"]):
        return f'<svg {props}><path d="M12 20h.01"/><path d="M2 8.82a15 15 0 0 1 20 0"/><path d="M5 12.859a10 10 0 0 1 14 0"/><path d="M8.5 16.429a5 5 0 0 1 7 0"/></svg>'
    if any(k in cat_norm for k in ["robot", "automa", "hardware", "eletron"]):
        return f'<svg {props}><path d="M12 8V4H8"/><rect width="16" height="12" x="4" y="8" rx="2"/><path d="M2 14h2"/><path d="M20 14h2"/><path d="M15 13v2"/><path d="M9 13v2"/></svg>'
    if any(k in cat_norm for k in ["varejo", "comerc", "consumo", "mercado", "negocio"]):
        return f'<svg {props}><path d="M6 2 3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4Z"/><path d="M3 6h18"/><path d="M16 10a4 4 0 0 1-8 0"/></svg>'

    return f'<svg {props}><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>'

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

# Configuração da Página
st.set_page_config(
    page_title="SPAI - Ambientes de Inovação SP",
    page_icon=CAMINHO_CTI if os.path.exists(CAMINHO_CTI) else "🏛️",
    layout="wide"
)

# Configuração para desativar aviso de tradução e definir o idioma como pt-BR
st.markdown("""
    <script>
        var doc = window.parent.document.documentElement;
        doc.setAttribute('lang', 'pt-BR');
        
        var head = window.parent.document.getElementsByTagName('head')[0];
        if (!window.parent.document.querySelector('meta[name="google"]')) {
            var meta = window.parent.document.createElement('meta');
            meta.name = 'google';
            meta.content = 'notranslate';
            head.appendChild(meta);
        }
    </script>
""", unsafe_allow_html=True)

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

# Estilização CSS Personalizada - Institucional SPAI (Alto Contraste e Elegância)
st.markdown("""
    <style>
    .block-container { padding-top: 1rem; padding-bottom: 0rem; }

    /* Card de Destaque Executivo - Total de Membros */
    .kpi-box-membros {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        color: #ffffff;
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
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-top: 6px;
        color: #94d2bd;
    }

    /* Cards Secundários do Topo */
    .kpi-box-top {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
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
    .kpi-lbl-top { font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.8px; margin-top: 6px; color: #1e293b; }

    /* Cards da Segunda Linha (Categorias) */
    .kpi-box-sec {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
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
    .kpi-lbl-sec { font-size: 10px; font-weight: 800; color: #1e293b; text-transform: uppercase; margin-top: 4px; line-height: 1.2; }
    .kpi-sub-sec { font-size: 9px; color: #475569; font-style: italic; margin-top: 2px; }

    /* Card Destaque - Filtro de Tipo de Membro */
    .filter-card-tipo {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border: 1px solid #94a3b8;
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
        color: #334155;
        margin-bottom: 10px;
    }

    div[data-testid="stPills"] button {
        border-radius: 8px !important;
        font-weight: 700 !important;
        font-size: 13px !important;
        padding: 8px 18px !important;
        border: 1px solid #64748b !important;
        transition: all 0.2s ease-in-out !important;
    }

    /* Passo 2: Legenda discreta e alinhada em linha única integrada aos botões */
    .st-key-cat_pills_vg div[data-testid="stPills"] {
        display: flex !important;
        flex-wrap: nowrap !important;
        overflow-x: auto !important;
        align-items: center !important;
        gap: 8px !important;
        padding-bottom: 4px !important;
    }
    .st-key-cat_pills_vg div[data-testid="stPills"] button {
        display: inline-flex !important;
        align-items: center !important;
        white-space: nowrap !important;
        border-radius: 999px !important;
        font-weight: 600 !important;
        font-size: 12.5px !important;
        padding: 6px 14px !important;
        border: 1px solid #cbd5e1 !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04) !important;
    }
    .st-key-cat_pills_vg div[data-testid="stPills"] button p {
        display: inline-flex !important;
        align-items: center !important;
        margin: 0 !important;
    }
    .st-key-cat_pills_vg div[data-testid="stPills"] button:hover {
        border-color: #94a3b8 !important;
        background-color: #f8fafc !important;
    }
    .st-key-cat_pills_vg div[data-testid="stPills"] button[aria-pressed="true"],
    .st-key-cat_pills_vg div[data-testid="stPills"] button[kind="pillsActive"] {
        background-color: #f1f5f9 !important;
        border: 1.5px solid #1B365D !important;
        color: #0f172a !important;
        font-weight: 700 !important;
        box-shadow: 0 1px 4px rgba(27, 54, 93, 0.14) !important;
    }

    /* Barra de Legenda Discreta do Mapa */
    .mapa-legenda-bar {
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 16px;
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 8px 14px;
        margin-bottom: 10px;
    }
    .mapa-legenda-item {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        font-size: 12px;
        font-weight: 600;
        color: #1e293b;
        white-space: nowrap;
    }

    .metric-card {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 12px;
        text-align: center;
        box-shadow: 0 1px 4px rgba(0,0,0,0.03);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .metric-val { font-size: 26px; font-weight: 800; color: #005f73; }
    .metric-lbl { font-size: 12px; color: #0f172a; font-weight: 700; }

    .concept-card {
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        border-left: 5px solid #005f73;
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 15px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    .concept-title {
        font-size: 15px;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 6px;
    }
    .concept-text {
        font-size: 13px;
        color: #1e293b;
        line-height: 1.5;
    }

    /* Grade de Cards de Distribuição (Verticais Tecnológicas e Segmentos de Mercado) */
    .dist-grid-top {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 16px;
        margin-bottom: 16px;
    }
    .dist-card-top {
        background: linear-gradient(90deg, #005f73 0%, #0a9396 100%);
        color: #ffffff;
        border-radius: 14px;
        padding: 22px 26px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        box-shadow: 0 4px 14px rgba(0, 95, 115, 0.22);
    }
    .dist-card-top-left {
        display: flex;
        align-items: center;
        gap: 16px;
    }
    .dist-card-top-icon {
        flex-shrink: 0;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    .dist-card-top-title {
        font-size: 17px;
        font-weight: 700;
        color: #ffffff;
        line-height: 1.3;
    }
    .dist-card-top-right {
        text-align: right;
        flex-shrink: 0;
    }
    .dist-card-top-pct {
        font-size: 34px;
        font-weight: 800;
        color: #ffffff;
        line-height: 1;
        letter-spacing: -0.5px;
    }
    .dist-card-top-sub {
        font-size: 11px;
        font-weight: 600;
        color: #94d2bd;
        margin-top: 5px;
    }

    .dist-grid-sub {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 16px;
    }
    .dist-card-sub {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 14px;
        padding: 18px 20px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        min-height: 136px;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
    }
    .dist-card-sub-hdr {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 10px;
    }
    .dist-card-sub-title {
        font-size: 14px;
        font-weight: 600;
        color: #0f172a;
        line-height: 1.35;
    }
    .dist-card-sub-icon {
        flex-shrink: 0;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    .dist-card-sub-ftr {
        margin-top: 16px;
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        gap: 8px;
    }
    .dist-card-sub-pct {
        font-size: 28px;
        font-weight: 800;
        color: #005f73;
        line-height: 1;
    }
    .dist-card-sub-qtd {
        font-size: 11px;
        font-weight: 600;
        color: #475569;
    }

    @media (max-width: 992px) {
        .dist-grid-top { grid-template-columns: 1fr; }
        .dist-grid-sub { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 576px) {
        .dist-grid-sub { grid-template-columns: 1fr; }
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

# Logo do CTI exibido no TOPO da barra lateral
if os.path.exists(CAMINHO_CTI):
    st.sidebar.image(CAMINHO_CTI, width='stretch')

# Logo Horizontal do SPAI (posicionado abaixo do logo do CTI)
if os.path.exists(CAMINHO_LOGO_HORIZONTAL):
    st.sidebar.image(CAMINHO_LOGO_HORIZONTAL, width='stretch')

painel_selecionado = st.sidebar.radio(
    "Navegação por Painel:",
    ["Visão Geral", "Verticais Tecnológicas", "Segmentos de Mercado"],
    index=0
)

st.sidebar.markdown("---")

# Link para o Mapeamento 2025
st.sidebar.markdown("📄 [**Acesse o Mapeamento de Ambientes de Inovação de São Paulo Ano-Base 2025**](https://drive.google.com/file/d/1cKcB8jkSQFSbfMoMK-bQAOCEo_NPZiX_/view?usp=sharing)")

st.sidebar.info("""
**Sistema Paulista de Ambientes de Inovação (SPAI)**

O **SPAI** é um programa da política pública de inovação do Estado de São Paulo, gerido pela **Secretaria de Ciência, Tecnologia e Inovação (SCTI-SP)**.

📌 **Objetivo:**
Incentivar investimentos em inovação, fortalecer a pesquisa científica e conectar parques tecnológicos, incubadoras, centros de inovação, centros de conexão e aceleradoras ao setor produtivo.

🔗 [Saiba mais no site oficial da SCTI](https://www.inovacao.sp.gov.br/sec_tecnologia_inovacao/programas/spai)
""")

# Caixa de texto "Colaboradores:" exibida logo acima do logo do Projeto Japi
st.sidebar.markdown("**Colaboradores:**")

if os.path.exists(CAMINHO_LOGO_JAPI):
    st.sidebar.image(CAMINHO_LOGO_JAPI, width='stretch')

# Função para renderização de painel fato com Grade de Cards (Top 2 + até 8 cards abaixo)
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
            st.markdown("<div style='text-align: center; font-size: 13px; font-weight: 700; color: #0f172a; margin-bottom: 4px;'>Categoria do Ambiente:</div>", unsafe_allow_html=True)
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
        # GRADE DE CARDS (TOP 2 EM DESTAQUE + ATÉ 8 CARDS ABAIXO = MÁXIMO 10)
        # ----------------------------------------------------
        st.subheader(f"📊 Distribuição por {titulo_item}")

        df_dist = df_filt_f.groupby(col_atuacao)[col_nome].nunique().reset_index()
        df_dist.columns = [titulo_item, "Quantidade de Ambientes"]
        df_dist = df_dist.sort_values(by="Quantidade de Ambientes", ascending=False).reset_index(drop=True)

        if not df_dist.empty:
            total_ocorrencias = df_dist["Quantidade de Ambientes"].sum()

            # Garante no máximo 10 cards (agrupando excedentes em "Outros" quando houver > 10)
            if len(df_dist) > 10:
                top9 = df_dist.iloc[:9].copy()
                outros_qtd = df_dist.iloc[9:]["Quantidade de Ambientes"].sum()
                outros_df = pd.DataFrame([{titulo_item: "Outros", "Quantidade de Ambientes": outros_qtd}])
                df_cards = pd.concat([top9, outros_df], ignore_index=True)
            else:
                df_cards = df_dist.head(10).copy()

            df_cards["Percentual"] = (df_cards["Quantidade de Ambientes"] / max(total_ocorrencias, 1)) * 100

            top2_df = df_cards.iloc[:2]
            restante_df = df_cards.iloc[2:10]

            # Primeira Linha: Top 2 Cards em Destaque
            html_cards = '<div class="dist-grid-top">'
            for _, row in top2_df.iterrows():
                nome_cat = str(row[titulo_item]).strip()
                qtd_cat = int(row["Quantidade de Ambientes"])
                pct_str = formatar_percentual_ptbr(row["Percentual"])
                icone = obter_icone_svg(nome_cat, cor="#ffffff", tamanho=34)
                lbl_amb = "ambiente" if qtd_cat == 1 else "ambientes"

                html_cards += (
                    f'<div class="dist-card-top">'
                    f'<div class="dist-card-top-left">'
                    f'<div class="dist-card-top-icon">{icone}</div>'
                    f'<div class="dist-card-top-title">{nome_cat}</div>'
                    f'</div>'
                    f'<div class="dist-card-top-right">'
                    f'<div class="dist-card-top-pct">{pct_str}</div>'
                    f'<div class="dist-card-top-sub">{qtd_cat} {lbl_amb}</div>'
                    f'</div>'
                    f'</div>'
                )
            html_cards += '</div>'

            # Linhas Abaixo: Demais Cards (até 8 cards em grade de 4 colunas)
            if not restante_df.empty:
                html_cards += '<div class="dist-grid-sub">'
                for _, row in restante_df.iterrows():
                    nome_cat = str(row[titulo_item]).strip()
                    qtd_cat = int(row["Quantidade de Ambientes"])
                    pct_str = formatar_percentual_ptbr(row["Percentual"])
                    icone = obter_icone_svg(nome_cat, cor="#005f73", tamanho=24)
                    lbl_amb = "ambiente" if qtd_cat == 1 else "ambientes"

                    html_cards += (
                        f'<div class="dist-card-sub">'
                        f'<div class="dist-card-sub-hdr">'
                        f'<div class="dist-card-sub-title">{nome_cat}</div>'
                        f'<div class="dist-card-sub-icon">{icone}</div>'
                        f'</div>'
                        f'<div class="dist-card-sub-ftr">'
                        f'<div class="dist-card-sub-pct">{pct_str}</div>'
                        f'<div class="dist-card-sub-qtd">{qtd_cat} {lbl_amb}</div>'
                        f'</div>'
                        f'</div>'
                    )
                html_cards += '</div>'

            st.markdown(html_cards, unsafe_allow_html=True)


# ====================================================
# PAINEL 1: VISÃO GERAL
# ====================================================
if painel_selecionado == "Visão Geral":
    # Exibição do Logo SPAI Vertical no tamanho máximo
    if os.path.exists(CAMINHO_LOGO_SPAI_VERTICAL):
        st.image(CAMINHO_LOGO_SPAI_VERTICAL, width='stretch')
    else:
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
    # FILTROS E SEGMENTADORES DE DADOS
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

    df_filt = df_spai.copy()
    if "Plenos" in str(tipo_sel_display) and df_filt.shape[1] > 4:
        df_filt = df_filt[df_filt.iloc[:, 4].astype(str).str.strip().str.title() == "Pleno"]
    elif "Associados" in str(tipo_sel_display) and df_filt.shape[1] > 4:
        df_filt = df_filt[df_filt.iloc[:, 4].astype(str).str.strip().str.title() == "Associado"]

    col_b_categorias = df_filt.iloc[:, 1].dropna().astype(str).str.strip().unique().tolist()
    cats_formatadas = list(set([c.title() for c in col_b_categorias if c != ""]))
    cats_ordenadas = ordenar_categorias_prioritarias(cats_formatadas)
    opcoes_categorias = ["Todas"] + cats_ordenadas

    # Injeta forma geométrica colorida (●, ◆, ■, ▲) ao lado do texto de cada botão do Passo 2
    css_legenda_pills = ["<style>"]
    for idx, opcao in enumerate(opcoes_categorias, start=1):
        if opcao == "Todas":
            continue
        estilo_cat = obter_estilo_categoria_mapa(opcao)
        cor_forma = estilo_cat["cor"]
        forma = estilo_cat["forma"]

        if forma == "circle":
            regra_forma = (
                f'content: "" !important; display: inline-block !important; width: 9px; height: 9px; min-width: 9px;'
                f'border-radius: 50%; background-color: {cor_forma}; margin-right: 7px; flex-shrink: 0;'
            )
        elif forma == "diamond":
            regra_forma = (
                f'content: "" !important; display: inline-block !important; width: 8px; height: 8px; min-width: 8px;'
                f'border-radius: 1px; transform: rotate(45deg); background-color: {cor_forma}; margin-right: 8px; flex-shrink: 0;'
            )
        elif forma == "square":
            regra_forma = (
                f'content: "" !important; display: inline-block !important; width: 9px; height: 9px; min-width: 9px;'
                f'border-radius: 1.5px; background-color: {cor_forma}; margin-right: 7px; flex-shrink: 0;'
            )
        elif forma == "triangle":
            regra_forma = (
                f'content: "" !important; display: inline-block !important; width: 10px; height: 9px; min-width: 10px;'
                f'clip-path: polygon(50% 0%, 0% 100%, 100% 100%); background-color: {cor_forma}; margin-right: 7px; flex-shrink: 0;'
            )
        else:
            regra_forma = (
                f'content: "" !important; display: inline-block !important; width: 8px; height: 8px; min-width: 8px;'
                f'border-radius: 50%; background-color: {cor_forma}; margin-right: 7px; flex-shrink: 0;'
            )

        css_legenda_pills.append(
            f'.st-key-cat_pills_vg div[data-testid="stPills"] button:nth-child({idx}) p::before, '
            f'.st-key-cat_pills_vg div[data-testid="stPills"] > div > div:nth-child({idx}) button p::before {{ {regra_forma} }}'
        )
    css_legenda_pills.append("</style>")
    st.markdown("\n".join(css_legenda_pills), unsafe_allow_html=True)

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
        # Legenda horizontal discreta com ícones geométricos alinhada acima do mapa
        itens_legenda_html = ['<div class="mapa-legenda-bar">']
        for cat_leg in cats_ordenadas:
            est_leg = obter_estilo_categoria_mapa(cat_leg)
            svg_leg = gerar_svg_forma_legenda(est_leg["forma"], est_leg["cor"])
            itens_legenda_html.append(f'<span class="mapa-legenda-item">{svg_leg}<span>{cat_leg}</span></span>')
        itens_legenda_html.append('</div>')
        st.markdown("".join(itens_legenda_html), unsafe_allow_html=True)

        df_m = df_filt.dropna(subset=["Lat", "Lon"]).copy()
        df_m["Lat"] = pd.to_numeric(df_m["Lat"], errors="coerce")
        df_m["Lon"] = pd.to_numeric(df_m["Lon"], errors="coerce")
        df_m = df_m.dropna(subset=["Lat", "Lon"]).reset_index(drop=True)

        if not df_m.empty:
            cores_pontos = []
            customdata_rows = []

            for _, r_m in df_m.iterrows():
                val_id_m = str(r_m.iloc[0]).strip() if df_m.shape[1] > 0 else ""
                val_cat_m = str(r_m.iloc[1]).strip().title() if df_m.shape[1] > 1 else ""
                val_reg_m = str(r_m.iloc[2]).strip() if df_m.shape[1] > 2 else ""
                val_mun_m = str(r_m.iloc[3]).strip() if df_m.shape[1] > 3 else ""
                val_tipo_m = str(r_m.iloc[4]).strip().title() if df_m.shape[1] > 4 else ""

                estilo_m = obter_estilo_categoria_mapa(val_cat_m)
                cores_pontos.append(estilo_m["cor"])
                customdata_rows.append([
                    val_id_m,
                    f"{estilo_m['simbolo']} {val_cat_m}" if val_cat_m else "",
                    val_reg_m,
                    val_mun_m,
                    val_tipo_m
                ])

            fig = go.Figure()

            # Camada de contorno branco fino (1px ao redor do marcador de 9px)
            fig.add_trace(go.Scattermap(
                lat=df_m["Lat"],
                lon=df_m["Lon"],
                mode="markers",
                marker=dict(
                    size=11,
                    color="#FFFFFF",
                    opacity=0.92
                ),
                hoverinfo="skip",
                showlegend=False
            ))

            # Marcadores minimalistas (9px, paleta institucional sóbria, 88% de opacidade)
            fig.add_trace(go.Scattermap(
                lat=df_m["Lat"],
                lon=df_m["Lon"],
                mode="markers",
                marker=dict(
                    size=9,
                    color=cores_pontos,
                    opacity=0.88
                ),
                text=df_m[col_nome_orig].astype(str).str.strip(),
                customdata=customdata_rows,
                hovertemplate=(
                    "<b>%{text}</b><br>"
                    "ID: %{customdata[0]}<br>"
                    "Categoria: %{customdata[1]}<br>"
                    "Região Administrativa: %{customdata[2]}<br>"
                    "Município: %{customdata[3]}<br>"
                    "Tipo: %{customdata[4]}"
                    "<extra></extra>"
                ),
                showlegend=False
            ))

            fig.update_layout(
                map=dict(
                    style="carto-positron",
                    zoom=6.1,
                    center=dict(lat=-22.3, lon=-48.8)
                ),
                margin=dict(r=0, t=0, l=0, b=0),
                height=450,
                showlegend=False
            )

            map_event = st.plotly_chart(
                fig,
                on_select="rerun",
                selection_mode="points",
                width='stretch'
            )

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
            height=450
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
