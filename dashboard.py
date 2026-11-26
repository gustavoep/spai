import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import os
import re
import json
import html
import unicodedata
from urllib.parse import quote
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

# Padroniza o rótulo de exibição da categoria no formato institucional
def formatar_categoria_exibicao(cat):
    if not isinstance(cat, str):
        cat = str(cat) if pd.notna(cat) else ""
    cat_limpa = cat.strip()
    if not cat_limpa:
        return ""
    c_norm = normalizar_texto(cat_limpa)
    if "parque" in c_norm:
        return "Parques Tecnológicos"
    if "centro de inov" in c_norm or "centros de inov" in c_norm:
        return "Centros de Inovação"
    if "incubadora" in c_norm:
        return "Incubadoras"
    if "hub" in c_norm or "conex" in c_norm:
        return "Centros de Conexão"
    if "aceleradora" in c_norm:
        return "Aceleradoras"
    return "Outros Formatos"

# Formata percentual no padrão brasileiro (ex: 31,5% ou 12%)
def formatar_percentual_ptbr(valor):
    arredondado = round(float(valor), 1)
    if arredondado == int(arredondado):
        return f"{int(arredondado)}%"
    return f"{arredondado:.1f}%".replace(".", ",")

# Retorna ícone SVG de linha simples para os cards de formato da Visão Geral
def obter_icone_formato_spai(tipo_formato, cor="#005f73", tamanho=24):
    props = f'width="{tamanho}" height="{tamanho}" viewBox="0 0 24 24" fill="none" stroke="{cor}" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"'
    if tipo_formato == "parque":
        return (
            f'<svg {props}>'
            f'<path d="M3 21h18"/>'
            f'<path d="M5 21V8l6-3v16"/>'
            f'<path d="M11 11h6v10"/>'
            f'<path d="M8 10h.01"/><path d="M8 14h.01"/><path d="M14 15h.01"/>'
            f'<circle cx="19" cy="9" r="2.5"/>'
            f'<path d="M19 11.5V21"/>'
            f'</svg>'
        )
    if tipo_formato == "incubadora":
        return (
            f'<svg {props}>'
            f'<rect x="4" y="3" width="16" height="18" rx="2"/>'
            f'<path d="M9 21v-4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v4"/>'
            f'<path d="M8 7h2"/><path d="M14 7h2"/>'
            f'<path d="M8 11h2"/><path d="M14 11h2"/>'
            f'</svg>'
        )
    if tipo_formato == "centro_inovacao":
        return (
            f'<svg {props}>'
            f'<circle cx="12" cy="12" r="3"/>'
            f'<circle cx="5" cy="6" r="2"/>'
            f'<circle cx="19" cy="6" r="2"/>'
            f'<circle cx="5" cy="18" r="2"/>'
            f'<circle cx="19" cy="18" r="2"/>'
            f'<path d="M6.7 7.3 9.8 10.2"/>'
            f'<path d="M17.3 7.3 14.2 10.2"/>'
            f'<path d="M6.7 16.7 9.8 13.8"/>'
            f'<path d="M17.3 16.7 14.2 13.8"/>'
            f'</svg>'
        )
    if tipo_formato == "distrito":
        return (
            f'<svg {props}>'
            f'<polygon points="3 6 9 3 15 6 21 3 21 18 15 21 9 18 3 21"/>'
            f'<line x1="9" y1="3" x2="9" y2="18"/>'
            f'<line x1="15" y1="6" x2="15" y2="21"/>'
            f'</svg>'
        )
    return ""

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

# Função para ordenar a lista de categorias na ordem institucional padronizada
def ordenar_categorias_prioritarias(lista_cats):
    ordem_padrao = {
        "Parques Tecnológicos": 0,
        "Centros de Inovação": 1,
        "Incubadoras": 2,
        "Centros de Conexão": 3,
        "Aceleradoras": 4,
        "Outros Formatos": 5,
    }
    return sorted(lista_cats, key=lambda c: (ordem_padrao.get(c, 999), c))

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

# URL ao vivo da planilha do Google Sheets — aba 'Membros SPAI'
SHEET_ID = "1yXVrdrAZW9QCYHIfp2MsyIa2iFsVAhGZYxnOst_UG_M"
ABA_PLANILHA = "Membros SPAI"
GOOGLE_SHEET_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq"
    f"?tqx=out:csv&headers=1&sheet={quote(ABA_PLANILHA)}"
)

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

# Estilização CSS Personalizada - Institucional SPAI
st.markdown("""
    <style>
    .block-container { padding-top: 0.75rem; padding-bottom: 0.75rem; }

    /* Card de Destaque Executivo Principal - Total de Membros */
    .kpi-box-membros {
        background: linear-gradient(135deg, #0f172a 0%, #005f73 100%);
        color: #ffffff;
        border-radius: 12px;
        padding: 15px 14px;
        text-align: center;
        box-shadow: 0 6px 18px rgba(0, 95, 115, 0.24);
        border: 1px solid #0a9396;
        border-left: 6px solid #94d2bd;
        min-height: 82px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    .kpi-val-membros {
        font-size: 38px;
        font-weight: 900;
        line-height: 1;
        color: #ffffff;
        letter-spacing: -0.8px;
    }
    .kpi-lbl-membros {
        font-size: 11px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.9px;
        margin-top: 6px;
        color: #94d2bd;
    }

    /* Cards Secundários do Topo */
    .kpi-box-top {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-top: 4px solid #005f73;
        border-radius: 10px;
        padding: 13px 8px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
        min-height: 82px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    .kpi-val-top { font-size: 28px; font-weight: 800; line-height: 1; color: #005f73; }
    .kpi-lbl-top { font-size: 10px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.7px; margin-top: 6px; color: #1e293b; }

    /* Grupo de Destaque dos 3 Formatos Principais (Parques, Centros de Inovação, Incubadoras) */
    .formato-destaque-card {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-top: 4px solid #005f73;
        border-radius: 12px;
        padding: 16px 20px 14px 20px;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        min-height: 136px;
        transition: border-color 0.15s ease, box-shadow 0.15s ease;
    }
    .formato-destaque-card:hover {
        border-color: #0a9396;
        box-shadow: 0 4px 14px rgba(0, 95, 115, 0.10);
    }
    .formato-destaque-top {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 12px;
    }
    .formato-destaque-val {
        font-size: 44px;
        font-weight: 800;
        color: #0f172a;
        line-height: 1;
        letter-spacing: -1px;
    }
    .formato-destaque-icon {
        width: 38px;
        height: 38px;
        border-radius: 9px;
        background-color: #f0fdfa;
        border: 1px solid #ccfbf1;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
    }
    .formato-destaque-lbl {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        font-size: 13.5px;
        font-weight: 700;
        color: #0f172a;
        margin-top: 8px;
        line-height: 1.25;
    }
    .formato-destaque-meta {
        margin-top: 10px;
    }
    .formato-destaque-prop {
        display: flex;
        align-items: center;
        justify-content: space-between;
        font-size: 11.5px;
        font-weight: 600;
        color: #475569;
        margin-bottom: 5px;
    }
    .formato-destaque-prop strong {
        color: #005f73;
        font-weight: 800;
    }
    .formato-progress-track {
        width: 100%;
        height: 5px;
        background-color: #e2e8f0;
        border-radius: 999px;
        overflow: hidden;
    }
    .formato-progress-fill {
        height: 100%;
        background: linear-gradient(90deg, #005f73 0%, #0a9396 100%);
        border-radius: 999px;
    }

    /* Faixa Secundária dos Demais Formatos (com destaque para Distritos de Inovação) */
    .formato-sec-bar {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr 1.35fr;
        gap: 10px;
        margin-top: 10px;
    }
    .formato-sec-item {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 9px;
        padding: 9px 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        min-height: 46px;
    }
    .formato-sec-item-distrito {
        background: linear-gradient(135deg, #f0fdfa 0%, #e6f4f1 100%);
        border: 1.5px solid #0a9396;
        border-left: 5px solid #005f73;
        border-radius: 10px;
        padding: 10px 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        min-height: 48px;
        box-shadow: 0 3px 10px rgba(0, 95, 115, 0.12);
    }
    .formato-sec-lbl {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        font-size: 12.5px;
        font-weight: 600;
        color: #334155;
        line-height: 1.25;
    }
    .formato-sec-lbl-distrito {
        display: inline-flex;
        align-items: center;
        gap: 7px;
        font-size: 13px;
        font-weight: 800;
        color: #005f73;
        line-height: 1.25;
    }
    .formato-sec-right {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-shrink: 0;
    }
    .formato-sec-val {
        font-size: 18px;
        font-weight: 800;
        color: #0f172a;
        line-height: 1;
    }
    .formato-sec-pct {
        font-size: 11px;
        font-weight: 600;
        color: #64748b;
    }
    .badge-em-formacao {
        display: inline-flex;
        align-items: center;
        background-color: #fef3c7;
        color: #92400e;
        border: 1px solid #f59e0b;
        font-size: 11px;
        font-weight: 800;
        padding: 4px 10px;
        border-radius: 999px;
        white-space: nowrap;
        letter-spacing: 0.25px;
        box-shadow: 0 1px 3px rgba(146, 64, 14, 0.12);
    }

    @media (max-width: 992px) {
        .formato-sec-bar { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 576px) {
        .formato-sec-bar { grid-template-columns: 1fr; }
    }

    /* Tooltip Interativo do Botão '?' (Tipo de Membro, Formatos e Distritos de Inovação) */
    .tipo-membro-hdr {
        position: relative;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        font-size: 13px;
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 6px;
        line-height: 1.4;
    }
    .help-tip-wrap {
        position: relative;
        display: inline-flex;
        align-items: center;
        justify-content: center;
    }
    .help-tip-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background-color: #e2e8f0;
        color: #0f172a;
        border: 1px solid #cbd5e1;
        border-radius: 50%;
        width: 17px;
        height: 17px;
        font-size: 11px;
        font-weight: 800;
        line-height: 1;
        cursor: help;
        transition: all 0.15s ease;
    }
    .help-tip-wrap:hover .help-tip-btn,
    .help-tip-wrap:focus-within .help-tip-btn {
        background-color: #005f73;
        color: #ffffff;
        border-color: #005f73;
    }
    .help-tip-popover {
        visibility: hidden;
        opacity: 0;
        position: absolute;
        left: 0;
        top: calc(100% + 8px);
        width: 410px;
        max-width: 88vw;
        background-color: #0f172a;
        color: #f8fafc;
        border: 1px solid #334155;
        border-left: 4px solid #0a9396;
        border-radius: 10px;
        padding: 12px 14px;
        font-size: 12px;
        font-weight: 400;
        line-height: 1.5;
        text-align: left;
        box-shadow: 0 10px 28px rgba(15, 23, 42, 0.28);
        z-index: 9999;
        pointer-events: none;
        transition: opacity 0.16s ease, visibility 0.16s ease;
    }
    .help-tip-popover.help-tip-popover-right {
        left: auto;
        right: 0;
    }
    .help-tip-popover p {
        margin: 0 0 8px 0;
        color: #f8fafc !important;
        font-size: 12px !important;
        line-height: 1.5 !important;
    }
    .help-tip-popover p:last-child {
        margin-bottom: 0;
    }
    .help-tip-popover strong {
        color: #94d2bd;
        font-weight: 700;
    }
    .help-tip-wrap:hover .help-tip-popover,
    .help-tip-wrap:focus-within .help-tip-popover {
        visibility: visible;
        opacity: 1;
    }

    /* Segmented Control (Nível 1 - Tipo de Membro) */
    div[data-testid="stSegmentedControl"] {
        background-color: #f1f5f9 !important;
        padding: 3px !important;
        border-radius: 8px !important;
        border: 1px solid #cbd5e1 !important;
        min-height: 40px !important;
        display: flex !important;
        align-items: center !important;
    }
    div[data-testid="stSegmentedControl"] button {
        border-radius: 6px !important;
        font-weight: 700 !important;
        font-size: 12.5px !important;
        color: #1e293b !important;
        min-height: 32px !important;
        transition: all 0.15s ease-in-out !important;
    }
    div[data-testid="stSegmentedControl"] button[aria-pressed="true"],
    div[data-testid="stSegmentedControl"] button[aria-checked="true"],
    div[data-testid="stSegmentedControl"] button[data-selected="true"] {
        background-color: #005f73 !important;
        color: #ffffff !important;
        box-shadow: 0 2px 6px rgba(0, 95, 115, 0.25) !important;
    }
    div[data-testid="stSegmentedControl"] button[aria-pressed="true"] p,
    div[data-testid="stSegmentedControl"] button[aria-checked="true"] p,
    div[data-testid="stSegmentedControl"] button[data-selected="true"] p {
        color: #ffffff !important;
    }

    /* Pills (Nível 2 - Subcategorias) */
    div[data-testid="stPills"] {
        margin-top: 2px !important;
        margin-bottom: 6px !important;
    }
    div[data-testid="stPills"] button {
        border-radius: 999px !important;
        font-weight: 600 !important;
        font-size: 12.5px !important;
        padding: 5px 14px !important;
        border: 1px solid #cbd5e1 !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
        transition: all 0.15s ease-in-out !important;
    }
    div[data-testid="stPills"] button:hover {
        border-color: #005f73 !important;
        color: #005f73 !important;
    }
    div[data-testid="stPills"] button[aria-pressed="true"],
    div[data-testid="stPills"] button[aria-checked="true"],
    div[data-testid="stPills"] button[data-selected="true"],
    div[data-testid="stPills"] button[kind="pillsActive"] {
        background-color: #005f73 !important;
        border-color: #005f73 !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        box-shadow: 0 2px 6px rgba(0, 95, 115, 0.25) !important;
    }
    div[data-testid="stPills"] button[aria-pressed="true"] p,
    div[data-testid="stPills"] button[aria-checked="true"] p,
    div[data-testid="stPills"] button[data-selected="true"] p,
    div[data-testid="stPills"] button[kind="pillsActive"] p {
        color: #ffffff !important;
    }

    /* Lista Interativa Compacta de Ambientes */
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stVerticalBlock"] {
        gap: 6px !important;
    }
    div[data-testid="stButton"] button:has(p code) {
        width: 100% !important;
        display: flex !important;
        align-items: center !important;
        justify-content: flex-start !important;
        text-align: left !important;
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-left: 4px solid #0a9396 !important;
        border-radius: 8px !important;
        padding: 8px 12px 8px 14px !important;
        min-height: 42px !important;
        height: auto !important;
        cursor: pointer !important;
        transition: all 0.14s ease !important;
    }
    div[data-testid="stButton"] button:has(p code) > div,
    div[data-testid="stButton"] button:has(p code) div[data-testid="stMarkdownContainer"] {
        width: 100% !important;
        flex: 1 1 100% !important;
        display: flex !important;
        align-items: center !important;
        justify-content: flex-start !important;
        text-align: left !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    div[data-testid="stButton"] button:has(p code) p {
        width: 100% !important;
        flex: 1 1 100% !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        gap: 12px !important;
        font-size: 12.5px !important;
        font-weight: 600 !important;
        line-height: 1.35 !important;
        color: #0f172a !important;
        margin: 0 !important;
        padding: 0 !important;
        text-align: left !important;
    }
    div[data-testid="stButton"] button:has(p code) p code {
        flex-shrink: 0 !important;
        margin-left: auto !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background-color: #e6f4f1 !important;
        color: #005f73 !important;
        border: 1px solid #94d2bd !important;
        border-radius: 999px !important;
        padding: 3px 11px !important;
        font-family: inherit !important;
        font-size: 11px !important;
        font-weight: 800 !important;
        letter-spacing: 0.2px !important;
        white-space: nowrap !important;
        transition: all 0.15s ease !important;
    }
    div[data-testid="stButton"] button:has(p code):hover {
        background-color: #f0fdfa !important;
        border-color: #005f73 !important;
        box-shadow: 0 2px 8px rgba(0, 95, 115, 0.14) !important;
    }
    div[data-testid="stButton"] button:has(p code):hover p {
        color: #005f73 !important;
    }
    div[data-testid="stButton"] button:has(p code):hover p code {
        background-color: #005f73 !important;
        color: #ffffff !important;
        border-color: #005f73 !important;
        transform: translateX(2px) !important;
    }

    /* Painel Lateral Sobreposto (Drawer) */
    .drawer-panel {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-top: 5px solid #005f73;
        border-radius: 12px;
        padding: 16px 18px;
        box-shadow: -6px 8px 24px rgba(15, 23, 42, 0.14);
        margin-bottom: 10px;
    }
    .drawer-badge-row {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-bottom: 10px;
    }
    .drawer-badge {
        background-color: #e6f4f1;
        color: #005f73;
        border: 1px solid #94d2bd;
        font-size: 11px;
        font-weight: 700;
        padding: 3px 9px;
        border-radius: 999px;
    }
    .drawer-badge-dark {
        background-color: #1e293b;
        color: #ffffff;
        font-size: 11px;
        font-weight: 700;
        padding: 3px 9px;
        border-radius: 999px;
    }
    .drawer-title {
        font-size: 16.5px;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.3;
        margin-bottom: 10px;
    }
    .drawer-section-lbl {
        font-size: 11px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        color: #475569;
        margin-top: 10px;
        margin-bottom: 3px;
    }
    .drawer-value {
        font-size: 12.5px;
        color: #0f172a;
        line-height: 1.45;
    }
    .drawer-subvalue {
        color: #475569;
        font-size: 11.5px;
        margin-top: 2px;
    }
    .drawer-resolucao-box {
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        border-left: 3.5px solid #005f73;
        border-radius: 7px;
        padding: 9px 11px;
        margin-top: 4px;
    }
    .drawer-resolucao-txt {
        font-size: 12px;
        font-weight: 700;
        color: #0f172a;
        line-height: 1.35;
    }
    .drawer-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 14px;
    }
    .drawer-btn-primary {
        display: inline-block;
        background-color: #005f73;
        color: #ffffff !important;
        text-decoration: none !important;
        font-size: 12px;
        font-weight: 700;
        padding: 8px 13px;
        border-radius: 8px;
    }
    .drawer-btn-primary:hover {
        background-color: #0a9396;
    }
    .drawer-btn-secondary {
        display: inline-block;
        background-color: #f1f5f9;
        color: #0f172a !important;
        border: 1px solid #cbd5e1;
        text-decoration: none !important;
        font-size: 12px;
        font-weight: 700;
        padding: 8px 13px;
        border-radius: 8px;
    }
    .drawer-btn-secondary:hover {
        background-color: #e2e8f0;
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

# Função de carregamento com cache da Base Geral (somente ambientes com dados completos)
@st.cache_data(ttl=120)
def carregar_base_geral():
    df = pd.read_csv(GOOGLE_SHEET_CSV_URL)

    qtd_cols_obrigatorias = min(6, df.shape[1])
    cols_obrigatorias = list(df.columns[:qtd_cols_obrigatorias])

    df = df.dropna(subset=cols_obrigatorias).copy()

    valores_invalidos = {"", "nan", "none", "null", "nat", "-"}
    for col in cols_obrigatorias:
        serie_limpa = df[col].astype(str).str.strip()
        df = df[~serie_limpa.str.lower().isin(valores_invalidos)]

    if df.shape[1] > 0:
        ids_numericos = pd.to_numeric(df.iloc[:, 0], errors="coerce")
        df = df[ids_numericos.notna()].copy()

    df = df.reset_index(drop=True)

    col_cidades = None
    for c in df.columns:
        if normalizar_texto(c) in {"cidades", "cidade", "municipios", "municipio"}:
            col_cidades = c
            break
    if col_cidades is None and df.shape[1] > 3:
        col_cidades = df.columns[3]

    cidade_col = df[col_cidades].astype(str).str.strip() if col_cidades is not None else pd.Series(dtype=str)
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

# Extrai campos opcionais extras da planilha (Endereço, Site, E-mail, Telefone, Resolução)
def extrair_campo_opcional(row, palavras_chave, somente_email_ou_fone=False):
    valores_invalidos = {"", "nan", "none", "null", "nat", "-"}
    for col in row.index:
        c_norm = normalizar_texto(col)
        if any(p in c_norm for p in palavras_chave):
            val = str(row[col]).strip() if pd.notna(row[col]) else ""
            if val.lower() not in valores_invalidos:
                if somente_email_ou_fone:
                    tem_email = "@" in val
                    tem_fone = bool(re.search(r"\d{4,}", val))
                    if not (tem_email or tem_fone):
                        continue
                return val
    return ""

# Extrai exclusivamente a descrição textual da Resolução de Credenciamento
def extrair_texto_resolucao(row):
    valores_invalidos = {"", "nan", "none", "null", "nat", "-"}
    for col in row.index:
        c_norm = normalizar_texto(col)
        if "resolucao" in c_norm and not any(k in c_norm for k in ["link", "url"]):
            val = str(row[col]).strip() if pd.notna(row[col]) else ""
            if val.lower() not in valores_invalidos:
                texto_limpo = re.sub(r"https?://[^\s<>\"']+", "", val).strip(" -–—|()")
                return texto_limpo
    return ""

# Monta o dicionário completo de informações do ambiente selecionado (mapa ou lista)
def montar_dados_ambiente(row_sel, col_id_orig, col_nome_orig, num_cols):
    nome_env = str(row_sel[col_nome_orig]).strip()
    id_env = int(pd.to_numeric(row_sel[col_id_orig], errors="coerce") or 0)
    cat_env = formatar_categoria_exibicao(row_sel.iloc[1]) if num_cols > 1 else ""
    reg_env = str(row_sel.iloc[2]).strip() if num_cols > 2 else ""
    mun_env = str(row_sel.iloc[3]).strip() if num_cols > 3 else ""
    tipo_env = str(row_sel.iloc[4]).strip().title() if num_cols > 4 else ""

    endereco_env = extrair_campo_opcional(row_sel, ["endereco", "logradouro", "rua", "localizacao", "cep"])
    site_env = extrair_campo_opcional(row_sel, ["site", "website", "url", "portal", "pagina"])
    contato_env = extrair_campo_opcional(row_sel, ["email", "e-mail", "telefone", "fone", "whatsapp", "contato"], somente_email_ou_fone=True)
    resolucao_txt = extrair_texto_resolucao(row_sel)

    v_env, s_env = buscar_detalhes_ambiente(nome_env)

    return {
        "nome": nome_env,
        "id": id_env,
        "categoria": cat_env,
        "regiao": reg_env,
        "municipio": mun_env,
        "tipo": tipo_env,
        "endereco": endereco_env,
        "site": site_env,
        "contato": contato_env,
        "resolucao_txt": resolucao_txt,
        "verticais": v_env,
        "segmentos": s_env,
    }

# Calcula o centro e o nível de zoom automático conforme os pontos filtrados
def calcular_enquadramento_mapa(df_pontos, filtro_regional_ativo=False):
    if df_pontos.empty:
        return {"lat": -22.3, "lon": -48.8}, 6.1

    lat_min, lat_max = df_pontos["Lat"].min(), df_pontos["Lat"].max()
    lon_min, lon_max = df_pontos["Lon"].min(), df_pontos["Lon"].max()

    centro = {
        "lat": float((lat_min + lat_max) / 2.0),
        "lon": float((lon_min + lon_max) / 2.0),
    }

    if not filtro_regional_ativo:
        return {"lat": -22.3, "lon": -48.8}, 6.1

    span = max(float(lat_max - lat_min), float(lon_max - lon_min))
    if span < 0.05:
        zoom = 11.2
    elif span < 0.35:
        zoom = 9.4
    elif span < 0.9:
        zoom = 8.4
    elif span < 1.8:
        zoom = 7.7
    else:
        zoom = 7.0

    return centro, zoom

# Renderiza o Painel Lateral Sobreposto (Drawer) com detalhes, resolução textual, endereço, site e contato
def renderizar_drawer_detalhes(dados):
    endereco_esc = html.escape(dados["endereco"]) if dados["endereco"] else f"{html.escape(dados['municipio'])} — RA de {html.escape(dados['regiao'])}, SP"
    regiao_esc = html.escape(dados["regiao"])
    nome_esc = html.escape(dados["nome"])
    tipo_esc = html.escape(dados["tipo"])
    cat_esc = html.escape(dados["categoria"])

    # Bloco da Resolução de Credenciamento (apenas descrição textual)
    res_txt = dados.get("resolucao_txt", "").strip()
    if res_txt:
        res_txt_esc = html.escape(res_txt)
        resolucao_block_html = (
            f'<div class="drawer-section-lbl">📜 Resolução de Credenciamento</div>'
            f'<div class="drawer-resolucao-box">'
            f'<div class="drawer-resolucao-txt">{res_txt_esc}</div>'
            f'</div>'
        )
    else:
        resolucao_block_html = (
            f'<div class="drawer-section-lbl">📜 Resolução de Credenciamento</div>'
            f'<div class="drawer-value"><i>Não informada na planilha.</i></div>'
        )

    # Link de Site
    site_raw = dados["site"].strip()
    if site_raw:
        site_url = site_raw if site_raw.startswith(("http://", "https://")) else f"https://{site_raw}"
    else:
        site_url = f"https://www.google.com/search?q={quote(dados['nome'] + ' ' + dados['municipio'])}"

    # Link de Contato
    contato_raw = dados["contato"].strip()
    if "@" in contato_raw:
        contato_url = f"mailto:{contato_raw}"
        contato_lbl = "✉️ Entrar em Contato"
    elif contato_raw:
        digitos_fone = re.sub(r"\D", "", contato_raw)
        contato_url = f"tel:+55{digitos_fone}" if digitos_fone else f"https://www.google.com/search?q={quote(dados['nome'] + ' contato')}"
        contato_lbl = f"📞 {html.escape(contato_raw)}"
    else:
        contato_url = f"https://www.google.com/search?q={quote(dados['nome'] + ' ' + dados['municipio'] + ' contato')}"
        contato_lbl = "✉️ Contato Institucional"

    verts_html = html.escape(", ".join(dados["verticais"])) if dados["verticais"] else "<i>Não mapeadas na fase atual.</i>"
    segs_html = html.escape(", ".join(dados["segmentos"])) if dados["segmentos"] else "<i>Não mapeados na fase atual.</i>"

    c_hdr, c_close = st.columns([5, 1])
    with c_hdr:
        st.markdown("<div style='font-size:12px; font-weight:800; color:#005f73; text-transform:uppercase; letter-spacing:0.6px; padding-top:4px;'>🏛️ Ficha do Ambiente</div>", unsafe_allow_html=True)
    with c_close:
        if st.button("✕", key="fechar_drawer_btn", help="Fechar painel de detalhes", use_container_width=True):
            st.session_state["ambiente_drawer_selecionado"] = None
            st.rerun()

    html_drawer = (
        f'<div class="drawer-panel">'
        f'<div class="drawer-badge-row">'
        f'<span class="drawer-badge-dark">ID #{dados["id"]}</span>'
        f'<span class="drawer-badge">{tipo_esc}</span>'
        f'<span class="drawer-badge">{cat_esc}</span>'
        f'</div>'
        f'<div class="drawer-title">{nome_esc}</div>'
        f'{resolucao_block_html}'
        f'<div class="drawer-section-lbl">📍 Endereço / Localização</div>'
        f'<div class="drawer-value">{endereco_esc}<div class="drawer-subvalue">Região Administrativa de {regiao_esc}</div></div>'
        f'<div class="drawer-section-lbl">💻 Verticais Tecnológicas</div>'
        f'<div class="drawer-value">{verts_html}</div>'
        f'<div class="drawer-section-lbl">🏢 Segmentos de Mercado</div>'
        f'<div class="drawer-value">{segs_html}</div>'
        f'<div class="drawer-actions">'
        f'<a href="{site_url}" target="_blank" rel="noopener noreferrer" class="drawer-btn-primary">🌐 Visitar Site</a>'
        f'<a href="{contato_url}" target="_blank" rel="noopener noreferrer" class="drawer-btn-secondary">{contato_lbl}</a>'
        f'</div>'
        f'</div>'
    )
    st.markdown(html_drawer, unsafe_allow_html=True)

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
        "Categoria": formatar_categoria_exibicao(r[col_spai_cat]) if col_spai_cat and pd.notna(r[col_spai_cat]) else "",
        "Regiao": str(r[col_spai_reg]).strip() if col_spai_reg and pd.notna(r[col_spai_reg]) else "",
        "Municipio": str(r[col_spai_mun]).strip() if col_spai_mun and pd.notna(r[col_spai_mun]) else "",
        "Tipo": str(r[col_spai_tipo]).strip().title() if col_spai_tipo and pd.notna(r[col_spai_tipo]) else ""
    }

# ----------------------------------------------------
# BARRA LATERAL - NAVEGAÇÃO ENTRE PAINÉIS
# ----------------------------------------------------

if os.path.exists(CAMINHO_CTI):
    st.sidebar.image(CAMINHO_CTI, width='stretch')

if os.path.exists(CAMINHO_LOGO_HORIZONTAL):
    st.sidebar.image(CAMINHO_LOGO_HORIZONTAL, width='stretch')

painel_selecionado = st.sidebar.radio(
    "Navegação por Painel:",
    ["Visão Geral", "Verticais Tecnológicas", "Segmentos de Mercado"],
    index=0
)

st.sidebar.markdown("---")

st.sidebar.markdown("📄 [**Acesse o Mapeamento de Ambientes de Inovação de São Paulo Ano-Base 2025**](https://drive.google.com/file/d/1cKcB8jkSQFSbfMoMK-bQAOCEo_NPZiX_/view?usp=sharing)")

st.sidebar.info("""
**Sistema Paulista de Ambientes de Inovação (SPAI)**

O **SPAI** é um programa da política pública de inovação do Estado de São Paulo, gerido pela **Secretaria de Ciência, Tecnologia e Inovação (SCTI-SP)**.

📌 **Objetivo:**
Incentivar investimentos em inovação, fortalecer a pesquisa científica e conectar parques tecnológicos, centros de inovação, incubadoras, centros de conexão e aceleradoras ao setor produtivo.

🔗 [Saiba mais no site oficial da SCTI](https://www.inovacao.sp.gov.br/sec_tecnologia_inovacao/programas/spai)
""")

st.sidebar.markdown("**Colaboradores:**")

if os.path.exists(CAMINHO_LOGO_JAPI):
    st.sidebar.image(CAMINHO_LOGO_JAPI, width='stretch')

# Função para renderização de painel fato com Grade de Cards
def renderizar_painel_fato(df_fato, titulo_item, item_plural, explicacao_conceitual):
    st.title(f"{item_plural}")
    st.caption(f"Mapeamento e distribuição de {item_plural.lower()} nos ambientes de inovação de São Paulo")

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

        st.subheader(f"📊 Distribuição por {titulo_item}")

        df_dist = df_filt_f.groupby(col_atuacao)[col_nome].nunique().reset_index()
        df_dist.columns = [titulo_item, "Quantidade de Ambientes"]
        df_dist = df_dist.sort_values(by="Quantidade de Ambientes", ascending=False).reset_index(drop=True)

        if not df_dist.empty:
            total_ocorrencias = df_dist["Quantidade de Ambientes"].sum()

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
    if os.path.exists(CAMINHO_LOGO_SPAI_VERTICAL):
        st.image(CAMINHO_LOGO_SPAI_VERTICAL, width='stretch')
    else:
        st.title("Visão Geral dos Ambientes de Inovação")
        st.caption("Sistema Paulista de Ambientes de Inovação — Secretaria de Ciência, Tecnologia e Inovação de SP")

    total_membros = len(df_spai)

    if df_spai.shape[1] > 4:
        tipo_col = df_spai.iloc[:, 4].astype(str).str.strip().str.title()
        plenos = (tipo_col == 'Pleno').sum()
        associados = (tipo_col == 'Associado').sum()
    else:
        plenos, associados = 0, 0

    col_cidades_vg = None
    for c in df_spai.columns:
        if normalizar_texto(c) in {"cidades", "cidade", "municipios", "municipio"}:
            col_cidades_vg = c
            break
    if col_cidades_vg is not None:
        municipios_count = df_spai[col_cidades_vg].dropna().astype(str).str.strip().nunique()
    elif df_spai.shape[1] > 3:
        municipios_count = df_spai.iloc[:, 3].dropna().astype(str).str.strip().nunique()
    else:
        municipios_count = 0

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

    top1, top2, top3, top4, top5 = st.columns([1.45, 1, 1, 1, 1])
    top1.markdown(f'''
        <div class="kpi-box-membros">
            <div class="kpi-val-membros">{total_membros}</div>
            <div class="kpi-lbl-membros">Total de Membros SPAI</div>
        </div>
    ''', unsafe_allow_html=True)
    top2.markdown(f'<div class="kpi-box-top"><div class="kpi-val-top">{plenos}</div><div class="kpi-lbl-top">Membros Plenos</div></div>', unsafe_allow_html=True)
    top3.markdown(f'<div class="kpi-box-top"><div class="kpi-val-top">{associados}</div><div class="kpi-lbl-top">Membros Associados</div></div>', unsafe_allow_html=True)
    top4.markdown(f'<div class="kpi-box-top"><div class="kpi-val-top">{municipios_count}</div><div class="kpi-lbl-top">Municípios</div></div>', unsafe_allow_html=True)
    top5.markdown(f'<div class="kpi-box-top"><div class="kpi-val-top">{regioes_count}</div><div class="kpi-lbl-top">Regiões Administrativas</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # Grupo de destaque: Parques Tecnológicos, Centros de Inovação e Incubadoras (nesta ordem)
    base_total = max(total_membros, 1)
    destaques_principais = [
        {
            "valor": parques,
            "label": "Parques Tecnológicos",
            "icone": obter_icone_formato_spai("parque", cor="#005f73", tamanho=22),
            "tooltip_html": (
                '<span class="help-tip-wrap">'
                '<span class="help-tip-btn" tabindex="0" aria-label="Explicação sobre Parques Tecnológicos">?</span>'
                '<div class="help-tip-popover" role="tooltip">'
                '<p><strong>Parques Tecnológicos</strong> são complexos de desenvolvimento econômico e tecnológico formalmente ligados a '
                'centros de excelência tecnológica, universidades e/ou centros de pesquisa, que visam fomentar economias baseadas no conhecimento '
                'por meio da integração entre pesquisa científica e tecnológica, empresas e organizações governamentais em um mesmo local físico.</p>'
                '<p><strong>Infraestrutura e serviços:</strong> além de prover espaço para negócios baseados em conhecimento e apoiar as inter-relações '
                'entre esses grupos, podem abrigar centros de pesquisa, desenvolvimento tecnológico, inovação, incubação, treinamento e prospecção, '
                'bem como servir de infraestrutura para feiras, exposições e desenvolvimento mercadológico.</p>'
                '</div>'
                '</span>'
            ),
        },
        {
            "valor": centros,
            "label": "Centros de Inovação",
            "icone": obter_icone_formato_spai("centro_inovacao", cor="#005f73", tamanho=22),
            "tooltip_html": (
                '<span class="help-tip-wrap">'
                '<span class="help-tip-btn" tabindex="0" aria-label="Explicação sobre Centros de Inovação Tecnológica">?</span>'
                '<div class="help-tip-popover" role="tooltip">'
                '<p><strong>Centro de Inovação Tecnológica</strong> é o empreendimento que concentra, integra e oferece um conjunto de mecanismos e serviços '
                'de suporte ao processo de inovação tecnológica das empresas, constituindo-se, também, em espaço de interação empresarial-acadêmica '
                'para o desenvolvimento de setores econômicos.</p>'
                '</div>'
                '</span>'
            ),
        },
        {
            "valor": incubadoras,
            "label": "Incubadoras",
            "icone": obter_icone_formato_spai("incubadora", cor="#005f73", tamanho=22),
            "tooltip_html": (
                '<span class="help-tip-wrap">'
                '<span class="help-tip-btn" tabindex="0" aria-label="Explicação sobre Incubadoras">?</span>'
                '<div class="help-tip-popover help-tip-popover-right" role="tooltip">'
                '<p><strong>Incubadoras</strong> são empreendimentos que, por tempo limitado, oferecem espaço físico para instalação de empresas e empreendimentos nascentes '
                'voltados ao desenvolvimento de produtos e processos intensivos em conhecimento, disponibilizando suporte gerencial e tecnológico, assim como outros serviços '
                'correlatos de valor agregado, com vista ao seu crescimento e consolidação.</p>'
                '</div>'
                '</span>'
            ),
        },
    ]

    dest_cols = st.columns(3)
    for col_d, item_d in zip(dest_cols, destaques_principais):
        qtd_d = int(item_d["valor"])
        pct_num = (qtd_d / base_total) * 100
        pct_fmt = formatar_percentual_ptbr(pct_num)
        largura_barra = min(max(pct_num, 0.0), 100.0)

        col_d.markdown(
            f'<div class="formato-destaque-card">'
            f'<div>'
            f'<div class="formato-destaque-top">'
            f'<div class="formato-destaque-val">{qtd_d}</div>'
            f'<div class="formato-destaque-icon">{item_d["icone"]}</div>'
            f'</div>'
            f'<div class="formato-destaque-lbl"><span>{item_d["label"]}</span>{item_d["tooltip_html"]}</div>'
            f'</div>'
            f'<div class="formato-destaque-meta">'
            f'<div class="formato-destaque-prop">'
            f'<span>{qtd_d} de {total_membros}</span>'
            f'<strong>{pct_fmt}</strong>'
            f'</div>'
            f'<div class="formato-progress-track">'
            f'<div class="formato-progress-fill" style="width: {largura_barra:.1f}%;"></div>'
            f'</div>'
            f'</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # Faixa secundária: formatos complementares e Distritos de Inovação em destaque
    pct_conexao = formatar_percentual_ptbr((centros_conexao / base_total) * 100)
    pct_acel = formatar_percentual_ptbr((aceleradoras / base_total) * 100)
    pct_outros = formatar_percentual_ptbr((outros / base_total) * 100)
    icone_distrito = obter_icone_formato_spai("distrito", cor="#005f73", tamanho=18)

    html_faixa_sec = (
        f'<div class="formato-sec-bar">'
        f'<div class="formato-sec-item">'
        f'<span class="formato-sec-lbl">Centros de Conexão</span>'
        f'<div class="formato-sec-right"><span class="formato-sec-val">{centros_conexao}</span><span class="formato-sec-pct">({pct_conexao})</span></div>'
        f'</div>'
        f'<div class="formato-sec-item">'
        f'<span class="formato-sec-lbl">Aceleradoras</span>'
        f'<div class="formato-sec-right"><span class="formato-sec-val">{aceleradoras}</span><span class="formato-sec-pct">({pct_acel})</span></div>'
        f'</div>'
        f'<div class="formato-sec-item">'
        f'<span class="formato-sec-lbl">Outros Formatos</span>'
        f'<div class="formato-sec-right"><span class="formato-sec-val">{outros}</span><span class="formato-sec-pct">({pct_outros})</span></div>'
        f'</div>'
        f'<div class="formato-sec-item-distrito">'
        f'<span class="formato-sec-lbl-distrito">'
        f'{icone_distrito}'
        f'<span>Distritos de Inovação</span>'
        f'<span class="help-tip-wrap">'
        f'<span class="help-tip-btn" tabindex="0" aria-label="Explicação sobre Distritos de Inovação">?</span>'
        f'<div class="help-tip-popover help-tip-popover-right" role="tooltip">'
        f'<p><strong>Distritos de Inovação</strong> são áreas urbanas planejadas que integram universidades, centros de pesquisa, '
        f'empresas e startups em um mesmo território conectado, estimulando a transferência de conhecimento, a revitalização urbana e a economia criativa.</p>'
        f'<p><strong>Por que ainda não há membros consolidados:</strong> trata-se de uma modalidade em estruturação territorial e institucional no Estado '
        f'(como a implantação do Distrito de Inovação de São Paulo na região da Cidade Universitária / Butantã), cujos critérios de credenciamento no SPAI ainda estão sendo regulamentados.</p>'
        f'</div>'
        f'</span>'
        f'</span>'
        f'<div class="formato-sec-right"><span class="badge-em-formacao">Em formação</span></div>'
        f'</div>'
        f'</div>'
    )
    st.markdown(html_faixa_sec, unsafe_allow_html=True)

    st.markdown("<hr style='margin: 16px 0 14px 0; border-color: #e2e8f0;'>", unsafe_allow_html=True)

    # ----------------------------------------------------
    # 1. BARRA DE FILTROS COMPACTA E ALINHADA
    # ----------------------------------------------------
    rotulo_todos = f"Todos ({total_membros})"
    rotulo_plenos = f"Plenos ({plenos})"
    rotulo_associados = f"Associados ({associados})"
    opcoes_tipo_display = [rotulo_todos, rotulo_plenos, rotulo_associados]

    if "tipo_membro_vg" not in st.session_state or st.session_state["tipo_membro_vg"] not in opcoes_tipo_display:
        st.session_state["tipo_membro_vg"] = rotulo_todos
    if "cat_pills_vg" not in st.session_state:
        st.session_state["cat_pills_vg"] = "Todas"
    if "regiao_sel_vg" not in st.session_state:
        st.session_state["regiao_sel_vg"] = "Todas"
    if "cidade_sel_vg" not in st.session_state:
        st.session_state["cidade_sel_vg"] = "Todas"
    if "busca_nome_vg" not in st.session_state:
        st.session_state["busca_nome_vg"] = ""
    if "ambiente_drawer_selecionado" not in st.session_state:
        st.session_state["ambiente_drawer_selecionado"] = None
    if "ultimo_ponto_mapa" not in st.session_state:
        st.session_state["ultimo_ponto_mapa"] = []

    def resetar_filtros_vg():
        st.session_state["tipo_membro_vg"] = rotulo_todos
        st.session_state["cat_pills_vg"] = "Todas"
        st.session_state["regiao_sel_vg"] = "Todas"
        st.session_state["cidade_sel_vg"] = "Todas"
        st.session_state["busca_nome_vg"] = ""
        st.session_state["ambiente_drawer_selecionado"] = None
        st.session_state["ultimo_ponto_mapa"] = []

    hdr_f, btn_f = st.columns([6, 1], vertical_alignment="center")
    with hdr_f:
        st.markdown("<div style='font-size:15.5px; font-weight:800; color:#0f172a;'>🔍 Explorar Ambientes de Inovação</div>", unsafe_allow_html=True)
    with btn_f:
        st.button("↺ Limpar filtros", key="btn_limpar_filtros_vg", on_click=resetar_filtros_vg, use_container_width=True)

    f_col1, f_col2, f_col3 = st.columns([2.2, 1.4, 1.4], vertical_alignment="bottom")

    with f_col1:
        html_tipo_hdr = (
            '<div class="tipo-membro-hdr">'
            '<span>Tipo de Membro</span>'
            '<span class="help-tip-wrap">'
            '<span class="help-tip-btn" tabindex="0" aria-label="Explicação sobre Tipo de Membro">?</span>'
            '<div class="help-tip-popover" role="tooltip">'
            '<p><strong>Membros plenos</strong> são ambientes de inovação geridos por organizações sem fins lucrativos ou públicas. '
            'Eles podem receber recursos do Governo do Estado por meio de convênios e termos de fomento.</p>'
            '<p><strong>Membros associados</strong> são ambientes de inovação geridos por organizações privadas com fins lucrativos. '
            'Embora não possam receber recursos orçamentários, eles podem participar em editais de financiamento e programas de capacitação.</p>'
            '</div>'
            '</span>'
            '</div>'
        )
        st.markdown(html_tipo_hdr, unsafe_allow_html=True)
        if hasattr(st, "segmented_control"):
            tipo_sel_display = st.segmented_control(
                label="Tipo de Membro",
                options=opcoes_tipo_display,
                key="tipo_membro_vg",
                label_visibility="collapsed",
            )
            if not tipo_sel_display:
                tipo_sel_display = rotulo_todos
        else:
            tipo_sel_display = st.radio(
                "Tipo de Membro",
                opcoes_tipo_display,
                key="tipo_membro_vg",
                horizontal=True,
                label_visibility="collapsed",
            )

    df_por_tipo = df_spai.copy()
    if "Plenos" in str(tipo_sel_display) and df_por_tipo.shape[1] > 4:
        df_por_tipo = df_por_tipo[df_por_tipo.iloc[:, 4].astype(str).str.strip().str.title() == "Pleno"]
    elif "Associados" in str(tipo_sel_display) and df_por_tipo.shape[1] > 4:
        df_por_tipo = df_por_tipo[df_por_tipo.iloc[:, 4].astype(str).str.strip().str.title() == "Associado"]

    cats_series = df_por_tipo.iloc[:, 1].apply(formatar_categoria_exibicao)
    contagem_cats = cats_series[cats_series != ""].value_counts().to_dict()
    cats_ordenadas = ordenar_categorias_prioritarias(list(contagem_cats.keys()))

    mapa_label_para_cat = {"Todas": "Todas"}
    opcoes_categorias_display = ["Todas"]
    for c in cats_ordenadas:
        qtd_c = contagem_cats.get(c, 0)
        lbl_c = f"{c} ({qtd_c})"
        opcoes_categorias_display.append(lbl_c)
        mapa_label_para_cat[lbl_c] = c

    cat_atual_limpa = mapa_label_para_cat.get(st.session_state["cat_pills_vg"], "Todas")
    if cat_atual_limpa == "Todas":
        st.session_state["cat_pills_vg"] = "Todas"
    else:
        correspondente = [lbl for lbl, raw in mapa_label_para_cat.items() if raw == cat_atual_limpa]
        st.session_state["cat_pills_vg"] = correspondente[0] if correspondente else "Todas"

    df_pos_cat = df_por_tipo.copy()
    if cat_atual_limpa != "Todas":
        df_pos_cat = df_pos_cat[df_pos_cat.iloc[:, 1].apply(formatar_categoria_exibicao) == cat_atual_limpa]

    with f_col2:
        regioes_unicas = ["Todas"] + sorted(list(df_pos_cat.iloc[:, 2].dropna().astype(str).str.strip().unique())) if df_pos_cat.shape[1] > 2 else ["Todas"]
        if st.session_state["regiao_sel_vg"] not in regioes_unicas:
            st.session_state["regiao_sel_vg"] = "Todas"
        regiao_sel = st.selectbox("Região Administrativa", regioes_unicas, key="regiao_sel_vg")

    df_pos_reg = df_pos_cat.copy()
    if regiao_sel != "Todas" and df_pos_reg.shape[1] > 2:
        df_pos_reg = df_pos_reg[df_pos_reg.iloc[:, 2].astype(str).str.strip() == regiao_sel]

    with f_col3:
        cidades_unicas = ["Todas"] + sorted(list(df_pos_reg.iloc[:, 3].dropna().astype(str).str.strip().unique())) if df_pos_reg.shape[1] > 3 else ["Todas"]
        if st.session_state["cidade_sel_vg"] not in cidades_unicas:
            st.session_state["cidade_sel_vg"] = "Todas"
        cidade_sel = st.selectbox("Município", cidades_unicas, key="cidade_sel_vg")

    st.markdown("<div style='font-size:12.5px; font-weight:700; color:#475569; margin: 10px 0 6px 0; line-height:1.4;'>Formato do Ambiente (filtrado pelo tipo acima):</div>", unsafe_allow_html=True)
    if hasattr(st, "pills"):
        cat_sel_display = st.pills(
            label="Formato do Ambiente",
            options=opcoes_categorias_display,
            key="cat_pills_vg",
            label_visibility="collapsed",
        )
        if not cat_sel_display:
            cat_sel_display = "Todas"
    else:
        cat_sel_display = st.selectbox(
            "Formato do Ambiente:",
            opcoes_categorias_display,
            key="cat_select_vg",
            label_visibility="collapsed",
        )

    cat_sel = mapa_label_para_cat.get(cat_sel_display, "Todas")

    df_filt = df_por_tipo.copy()
    if cat_sel != "Todas":
        df_filt = df_filt[df_filt.iloc[:, 1].apply(formatar_categoria_exibicao) == cat_sel]
    if regiao_sel != "Todas" and df_filt.shape[1] > 2:
        df_filt = df_filt[df_filt.iloc[:, 2].astype(str).str.strip() == regiao_sel]
    if cidade_sel != "Todas" and df_filt.shape[1] > 3:
        df_filt = df_filt[df_filt.iloc[:, 3].astype(str).str.strip() == cidade_sel]

    col_id_orig = df_filt.columns[0]
    col_nome_orig = df_filt.columns[5] if df_filt.shape[1] > 5 else df_filt.columns[0]

    termo_busca = st.session_state.get("busca_nome_vg", "").strip()
    if termo_busca:
        termo_norm = normalizar_texto(termo_busca)
        df_filt = df_filt[df_filt[col_nome_orig].astype(str).apply(lambda x: termo_norm in normalizar_texto(x))]

    df_filt = df_filt.reset_index(drop=True)

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    # ----------------------------------------------------
    # 2. ÁREA PRINCIPAL: MAPA + LISTA (OU DRAWER LATERAL)
    # ----------------------------------------------------
    drawer_ativo = st.session_state.get("ambiente_drawer_selecionado")
    if drawer_ativo is not None:
        col_m, col_l, col_drawer = st.columns([5.2, 3.3, 3.5], vertical_alignment="top")
    else:
        col_m, col_l = st.columns([6.2, 3.8], vertical_alignment="top")
        col_drawer = None

    map_event = None
    df_m = pd.DataFrame()

    with col_m:
        df_m = df_filt.dropna(subset=["Lat", "Lon"]).copy()
        df_m["Lat"] = pd.to_numeric(df_m["Lat"], errors="coerce")
        df_m["Lon"] = pd.to_numeric(df_m["Lon"], errors="coerce")
        df_m = df_m.dropna(subset=["Lat", "Lon"]).reset_index(drop=True)

        if not df_m.empty:
            df_m["_Categoria_Fmt"] = df_m.iloc[:, 1].apply(formatar_categoria_exibicao) if df_m.shape[1] > 1 else ""
            df_m["_Regiao_Fmt"] = df_m.iloc[:, 2].astype(str).str.strip() if df_m.shape[1] > 2 else ""
            df_m["_Municipio_Fmt"] = df_m.iloc[:, 3].astype(str).str.strip() if df_m.shape[1] > 3 else ""
            df_m["_Tipo_Fmt"] = df_m.iloc[:, 4].astype(str).str.strip().str.title() if df_m.shape[1] > 4 else ""

            filtro_reg_ativo = (regiao_sel != "Todas") or (cidade_sel != "Todas") or bool(termo_busca)
            centro_mapa, zoom_mapa = calcular_enquadramento_mapa(df_m, filtro_regional_ativo=filtro_reg_ativo)

            hover_dict = {
                "Lat": False,
                "Lon": False,
                "_Categoria_Fmt": True,
                "_Municipio_Fmt": True,
                "_Regiao_Fmt": True,
                "_Tipo_Fmt": True,
            }
            labels_dict = {
                "_Categoria_Fmt": "Categoria",
                "_Municipio_Fmt": "Município",
                "_Regiao_Fmt": "Região Administrativa",
                "_Tipo_Fmt": "Tipo",
            }

            fig = px.scatter_map(
                df_m,
                lat="Lat",
                lon="Lon",
                hover_name=col_nome_orig,
                hover_data=hover_dict,
                labels=labels_dict,
                zoom=zoom_mapa,
                center=centro_mapa,
                height=512,
            )
            fig.update_layout(
                map_style="carto-positron",
                margin={"r": 0, "t": 0, "l": 0, "b": 0},
                uirevision=f"{regiao_sel}_{cidade_sel}_{termo_busca}",
            )

            nome_ativo = drawer_ativo["nome"] if drawer_ativo else ""
            cores_pinos = [
                "#e85d04" if str(n).strip() == nome_ativo else "#005f73"
                for n in df_m[col_nome_orig]
            ]
            tamanhos_pinos = [
                19 if str(n).strip() == nome_ativo else 13
                for n in df_m[col_nome_orig]
            ]
            fig.update_traces(marker=dict(size=tamanhos_pinos, color=cores_pinos, opacity=0.9))

            map_event = st.plotly_chart(
                fig,
                on_select="rerun",
                selection_mode="points",
                key="mapa_spai_vg",
                use_container_width=True,
            )
        else:
            st.warning("⚠️ Nenhum ambiente com coordenadas mapeadas encontrado para os filtros aplicados.")

    with col_l:
        st.text_input(
            label="Buscar ambiente",
            placeholder="🔎 Buscar ambiente por nome (ex: SUPERA, Inova)...",
            key="busca_nome_vg",
            label_visibility="collapsed",
        )

        qtd_encontrados = len(df_filt)
        lbl_resultados = "ambiente encontrado" if qtd_encontrados == 1 else "ambientes encontrados"
        st.markdown(
            f"<div style='font-size:12px; font-weight:700; color:#475569; margin: 2px 0 8px 2px;'>"
            f"{qtd_encontrados} {lbl_resultados} · Clique para abrir detalhes"
            f"</div>",
            unsafe_allow_html=True,
        )

        with st.container(height=440, border=True):
            if df_filt.empty:
                st.info("Nenhum ambiente corresponde à busca atual.")
            else:
                for idx_row, row_item in df_filt.iterrows():
                    nome_item = str(row_item[col_nome_orig]).strip().replace("`", "")
                    mun_item = str(row_item.iloc[3]).strip().replace("`", "") if df_filt.shape[1] > 3 else ""
                    rotulo_botao = f"{nome_item} · {mun_item} `Ver ficha ➔`" if mun_item else f"{nome_item} `Ver ficha ➔`"

                    if st.button(rotulo_botao, key=f"env_btn_{idx_row}_{nome_item}", use_container_width=True):
                        st.session_state["ambiente_drawer_selecionado"] = montar_dados_ambiente(
                            row_item,
                            col_id_orig,
                            col_nome_orig,
                            df_filt.shape[1],
                        )
                        st.rerun()

    pontos_mapa = []
    if map_event and isinstance(map_event, dict) and "selection" in map_event:
        pontos_mapa = map_event["selection"].get("points", [])

    if pontos_mapa != st.session_state["ultimo_ponto_mapa"]:
        st.session_state["ultimo_ponto_mapa"] = list(pontos_mapa)
        if pontos_mapa and not df_m.empty:
            idx_ponto = pontos_mapa[0].get("point_index")
            if idx_ponto is not None and 0 <= idx_ponto < len(df_m):
                st.session_state["ambiente_drawer_selecionado"] = montar_dados_ambiente(
                    df_m.iloc[idx_ponto],
                    col_id_orig,
                    col_nome_orig,
                    df_filt.shape[1],
                )
                st.rerun()

    if col_drawer is not None and st.session_state.get("ambiente_drawer_selecionado") is not None:
        with col_drawer:
            renderizar_drawer_detalhes(st.session_state["ambiente_drawer_selecionado"])

    # ----------------------------------------------------
    # 3. HOVER SINCRONIZADO LISTA -> DESTAQUE NO PINO DO MAPA
    # ----------------------------------------------------
    if not df_m.empty:
        nomes_mapa_json = json.dumps(df_m[col_nome_orig].astype(str).str.strip().tolist(), ensure_ascii=False)
        nome_selecionado_json = json.dumps(drawer_ativo["nome"] if drawer_ativo else "", ensure_ascii=False)
        components.html(
            f"""
            <script>
            (function() {{
                const doc = window.parent.document;
                const nomesMapa = {nomes_mapa_json};
                const nomeAtivo = {nome_selecionado_json};

                function destacarPinoMapa(nomeHover) {{
                    const plots = doc.querySelectorAll('.js-plotly-plot');
                    plots.forEach(gd => {{
                        if (!gd || !gd.data || !gd.data.length) return;
                        const trace = gd.data[0];
                        if (!trace || !trace.hovertext) return;

                        const cores = nomesMapa.map(n => {{
                            if (nomeHover && n === nomeHover) return '#e85d04';
                            if (nomeAtivo && n === nomeAtivo) return '#e85d04';
                            return '#005f73';
                        }});
                        const tamanhos = nomesMapa.map(n => {{
                            if (nomeHover && n === nomeHover) return 21;
                            if (nomeAtivo && n === nomeAtivo) return 19;
                            return 13;
                        }});

                        if (window.parent.Plotly) {{
                            window.parent.Plotly.restyle(gd, {{
                                'marker.color': [cores],
                                'marker.size': [tamanhos]
                            }}, [0]);
                        }}
                    }});
                }}

                function vincularHoverLista() {{
                    const botoes = doc.querySelectorAll('div[data-testid="stButton"] button');
                    botoes.forEach(btn => {{
                        if (!btn.querySelector('p code')) return;
                        if (btn.dataset.hoverBound === '1') return;
                        btn.dataset.hoverBound = '1';
                        const pEl = btn.querySelector('p');
                        const textoCompleto = ((pEl ? pEl.innerText : btn.innerText) || '').trim();
                        const nomeEnv = textoCompleto.split(' · ')[0].replace(/Ver ficha.*$/, '').trim();
                        btn.addEventListener('mouseenter', () => destacarPinoMapa(nomeEnv));
                        btn.addEventListener('mouseleave', () => destacarPinoMapa(null));
                    }});
                }}

                vincularHoverLista();
                const obs = new MutationObserver(vincularHoverLista);
                obs.observe(doc.body, {{ childList: true, subtree: true }});
            }})();
            </script>
            """,
            height=0,
            width=0,
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
