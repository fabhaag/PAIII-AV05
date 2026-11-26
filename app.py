# --------------------------------------------------
# Rodar com o COMANDO: py -m streamlit run app.py
# --------------------------------------------------

import os
import io
import base64
import shutil
import sqlite3
import zipfile
import warnings
import datetime
from urllib.parse import urlencode

import numpy as np
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing

# Configuração inicial da página Streamlit
st.set_page_config(
    page_title="Gerador de Certificados CMM - Schulz",
    page_icon="📏",
    layout="wide",
    initial_sidebar_state="expanded"
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.txt")

STRICT_NS = "http://openxmlformats.org"
TRANS_NS = "http://openxmlformats.org"


# ==========================================
# ESTILIZAÇÃO VISUAL (CSS CUSTOMIZADO)
# ==========================================
def aplicar_estilo_customizado():
    st.markdown(
        """
        <style>
        /* Fundo geral mais limpo */
        .stApp {
            background-color: #f5f7fa;
        }

        /* Banner superior */
        .hero-banner {
            background: linear-gradient(135deg, #002b5c 0%, #0052cc 100%);
            padding: 24px 28px;
            border-radius: 14px;
            color: white;
            box-shadow: 0 4px 16px rgba(0, 43, 92, 0.15);
            margin-bottom: 20px;
        }
        .hero-title {
            font-size: 28px;
            font-weight: 700;
            margin: 0 0 6px 0;
            color: #ffffff;
        }
        .hero-subtitle {
            font-size: 14px;
            color: #dbeafe;
            margin: 0 0 12px 0;
            line-height: 1.45;
        }
        .hero-badge {
            display: inline-block;
            background-color: rgba(255, 255, 255, 0.16);
            border: 1px solid rgba(255, 255, 255, 0.28);
            color: #ffffff;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            margin-right: 6px;
        }

        /* Imagem da MMC no cabeçalho */
        .mmc-img-wrapper {
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.22);
            border: 2px solid rgba(255, 255, 255, 0.25);
            background-color: #ffffff;
        }
        .mmc-img-wrapper img {
            width: 100%;
            height: 175px;
            object-fit: cover;
            display: block;
        }

        /* Cards de KPI */
        .kpi-card {
            background-color: #ffffff;
            border-radius: 12px;
            padding: 16px 18px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
            text-align: center;
        }
        .kpi-label {
            font-size: 12px;
            font-weight: 600;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
        }
        .kpi-value {
            font-size: 24px;
            font-weight: 700;
            color: #0f172a;
        }
        .kpi-aprovado {
            background-color: #ecfdf5;
            border: 1px solid #86efac;
        }
        .kpi-aprovado .kpi-value {
            color: #15803d;
        }
        .kpi-reprovado {
            background-color: #fef2f2;
            border: 1px solid #fca5a5;
        }
        .kpi-reprovado .kpi-value {
            color: #b91c1c;
        }

        /* Títulos de seção */
        .section-title {
            font-size: 18px;
            font-weight: 700;
            color: #002b5c;
            margin-bottom: 10px;
            border-left: 4px solid #0052cc;
            padding-left: 10px;
        }

        /* Botão principal de download */
        div.stDownloadButton > button {
            background: linear-gradient(90deg, #003d99 0%, #0052cc 100%);
            color: white;
            font-weight: 600;
            border: none;
            border-radius: 10px;
            padding: 0.65rem 1.2rem;
            box-shadow: 0 4px 10px rgba(0, 82, 204, 0.25);
        }
        div.stDownloadButton > button:hover {
            background: linear-gradient(90deg, #002b6b 0%, #0041a3 100%);
            color: white;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def carregar_imagem_base64(caminho_relativo):
    """Carrega uma imagem local e retorna em base64 para embutir em HTML."""
    caminho_completo = os.path.join(BASE_DIR, caminho_relativo)
    if not os.path.exists(caminho_completo):
        base, ext = os.path.splitext(caminho_completo)
        alternativo = base + ext.lower()
        if os.path.exists(alternativo):
            caminho_completo = alternativo
        else:
            return None
    with open(caminho_completo, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def carregar_configuracao():
    """Lê (ou cria) o arquivo config.txt e garante que as pastas existam."""
    config = {
        "PASTA_CERTIFICADOS": "certificados",
        "PASTA_DADOS": "dados"
    }

    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            f.write("# Configuração de diretórios do Gerador de Certificados CMM - Schulz\n")
            f.write("# Use caminhos relativos (ex: certificados) ou absolutos (ex: C:\\Schulz\\Certificados)\n")
            f.write("PASTA_CERTIFICADOS=certificados\n")
            f.write("PASTA_DADOS=dados\n")
    else:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            for linha in f:
                linha = linha.strip()
                if not linha or linha.startswith("#") or "=" not in linha:
                    continue
                chave, valor = linha.split("=", 1)
                chave = chave.strip().upper()
                valor = valor.strip()
                if chave in config and valor:
                    config[chave] = valor

    pasta_certificados = config["PASTA_CERTIFICADOS"]
    if not os.path.isabs(pasta_certificados):
        pasta_certificados = os.path.join(BASE_DIR, pasta_certificados)

    pasta_dados = config["PASTA_DADOS"]
    if not os.path.isabs(pasta_dados):
        pasta_dados = os.path.join(BASE_DIR, pasta_dados)

    os.makedirs(pasta_certificados, exist_ok=True)
    os.makedirs(pasta_dados, exist_ok=True)

    return pasta_certificados, pasta_dados


def extrair_ano_e_data(data_raw):
    """Normaliza a data para DD/MM/AAAA e retorna (ano, data_formatada)."""
    data_raw = str(data_raw).strip()
    if "-" in data_raw and len(data_raw.split("-")[0]) == 4:
        partes = data_raw.split("-")
        ano = partes[0]
        data_formatada = f"{partes[2]}/{partes[1]}/{partes[0]}"
    elif "/" in data_raw and len(data_raw.split("/")[-1]) == 4:
        partes = data_raw.split("/")
        ano = partes[-1]
        data_formatada = data_raw
    else:
        ano = datetime.datetime.now().strftime("%Y")
        data_formatada = data_raw.replace("-", "/")
    return ano, data_formatada


def inicializar_banco_sqlite(caminho_db):
    """Cria as tabelas no SQLite espelhando o modelo relacional do PostgreSQL."""
    with sqlite3.connect(caminho_db) as conn:
        conn.execute("PRAGMA foreign_keys = ON;")
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS metadados (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                dispositivo_controle TEXT,
                data_medicao TEXT,
                hora_medicao TEXT,
                ordem TEXT,
                desenho TEXT,
                numero_peca TEXT,
                cliente TEXT,
                metrologista TEXT,
                mmc TEXT,
                fornecedor TEXT,
                temperatura REAL
            );

            CREATE TABLE IF NOT EXISTS certificado (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_metadados INTEGER NOT NULL,
                fator_k REAL,
                incerteza_u REAL,
                status_final TEXT,
                data_emissao TEXT,
                FOREIGN KEY (id_metadados) REFERENCES metadados(id)
            );

            CREATE TABLE IF NOT EXISTS parametros (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_certificado INTEGER NOT NULL,
                caracteristica TEXT,
                nominal REAL,
                tolerancia_superior REAL,
                tolerancia_inferior REAL,
                medida_1 REAL,
                medida_2 REAL,
                media REAL,
                desvio_padrao REAL,
                erro REAL,
                aprovado INTEGER,
                FOREIGN KEY (id_certificado) REFERENCES certificado(id)
            );
        """)


def salvar_dados_sqlite(dados_tela, df_final, val_k, val_u, calibracao_sucesso, pasta_dados):
    """Persiste metadados, certificado e parâmetros no SQLite anual (calibracoes_{ano}.db)."""
    ano, _ = extrair_ano_e_data(dados_tela.get("Data", ""))
    caminho_db = os.path.join(pasta_dados, f"calibracoes_{ano}.db")
    inicializar_banco_sqlite(caminho_db)

    temp_raw = str(dados_tela.get("Temperatura", "")).replace(",", ".").strip()
    try:
        temp_val = float(temp_raw) if temp_raw else None
    except ValueError:
        temp_val = None

    status_final = "APROVADO" if calibracao_sucesso else "REPROVADO"
    data_emissao = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with sqlite3.connect(caminho_db) as conn:
        conn.execute("PRAGMA foreign_keys = ON;")
        cursor = conn.cursor()

        cursor.execute(
            """
            INSERT INTO metadados (
                dispositivo_controle, data_medicao, hora_medicao, ordem,
                desenho, numero_peca, cliente, metrologista, mmc, fornecedor, temperatura
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                dados_tela.get("Dispositivo de Controle", ""),
                dados_tela.get("Data", ""),
                dados_tela.get("Hora", ""),
                dados_tela.get("Ordem", ""),
                dados_tela.get("Desenho", ""),
                dados_tela.get("Numero peça", ""),
                dados_tela.get("Cliente", ""),
                dados_tela.get("Metrologista", ""),
                dados_tela.get("MMC", ""),
                dados_tela.get("Fornecedor", ""),
                temp_val,
            ),
        )
        id_metadados = cursor.lastrowid

        cursor.execute(
            """
            INSERT INTO certificado (id_metadados, fator_k, incerteza_u, status_final, data_emissao)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                id_metadados,
                round(float(val_k), 2),
                round(float(val_u), 4),
                status_final,
                data_emissao,
            ),
        )
        id_certificado = cursor.lastrowid

        registros_parametros = []
        for _, row in df_final.iterrows():
            nom = float(row["Nominal"]) if pd.notna(row["Nominal"]) else None
            tsup = float(row["Upper Tol"]) if pd.notna(row["Upper Tol"]) else None
            tinf = float(row["Lower Tol"]) if pd.notna(row["Lower Tol"]) else None
            m1 = float(row["Medida 1"]) if pd.notna(row["Medida 1"]) else None
            m2 = float(row["Medida 2"]) if pd.notna(row["Medida 2"]) else None
            media = float(row["Média"]) if pd.notna(row["Média"]) else None
            desvio = float(row["Desvio padrão"]) if pd.notna(row["Desvio padrão"]) else 0.0
            erro = abs(media - nom) if (media is not None and nom is not None) else None
            aprovado = 1 if bool(row["Aprovado"]) else 0

            registros_parametros.append((
                id_certificado,
                str(row["Characteristic"]),
                nom, tsup, tinf, m1, m2, media, desvio, erro, aprovado
            ))

        cursor.executemany(
            """
            INSERT INTO parametros (
                id_certificado, caracteristica, nominal, tolerancia_superior,
                tolerancia_inferior, medida_1, medida_2, media, desvio_padrao, erro, aprovado
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            registros_parametros,
        )

    return caminho_db


def exportar_sqlite_para_excel(ano, pasta_dados):
    """Exporta as 3 tabelas do SQLite do ano selecionado para Calibrações-{ano}.xlsx."""
    caminho_db = os.path.join(pasta_dados, f"calibracoes_{ano}.db")
    nome_excel = f"Calibrações-{ano}.xlsx"
    caminho_excel = os.path.join(pasta_dados, nome_excel)
    caminho_tmp = os.path.join(pasta_dados, f".Calibrações-{ano}.tmp.xlsx")

    with sqlite3.connect(caminho_db) as conn:
        df_meta = pd.read_sql_query("SELECT * FROM metadados", conn)
        df_cert = pd.read_sql_query("SELECT * FROM certificado", conn)
        df_param = pd.read_sql_query("SELECT * FROM parametros", conn)

    if "aprovado" in df_param.columns:
        df_param["aprovado"] = df_param["aprovado"].astype(bool)

    with pd.ExcelWriter(caminho_tmp, engine="openpyxl") as writer:
        df_meta.to_excel(writer, sheet_name="metadados", index=False)
        df_cert.to_excel(writer, sheet_name="certificado", index=False)
        df_param.to_excel(writer, sheet_name="parametros", index=False)

    os.replace(caminho_tmp, caminho_excel)

    with open(caminho_excel, "rb") as f:
        excel_bytes = f.read()

    return caminho_excel, nome_excel, excel_bytes


def salvar_pdf_com_versionamento(pdf_bytes, nome_arquivo, pasta_certificados):
    """
    Salva o PDF em pasta_certificados.
    Se já existir um PDF com o mesmo nome, move o antigo para certificados/historico/ com timestamp.
    """
    caminho_pdf = os.path.join(pasta_certificados, nome_arquivo)
    caminho_historico = None

    if os.path.exists(caminho_pdf):
        pasta_historico = os.path.join(pasta_certificados, "historico")
        os.makedirs(pasta_historico, exist_ok=True)

        base, ext = os.path.splitext(nome_arquivo)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_antigo = f"{base}_{timestamp}{ext}"
        caminho_historico = os.path.join(pasta_historico, nome_antigo)
        shutil.move(caminho_pdf, caminho_historico)

    with open(caminho_pdf, "wb") as f:
        f.write(pdf_bytes)

    return caminho_pdf, caminho_historico


def salvar_certificado_e_dados(pdf_bytes, nome_arquivo, dados_tela, df_final, val_k, val_u, calibracao_sucesso):
    """Callback executado ao clicar em baixar: versiona/salva o PDF e grava no SQLite."""
    pasta_certificados, pasta_dados = carregar_configuracao()

    try:
        caminho_pdf, caminho_historico = salvar_pdf_com_versionamento(
            pdf_bytes, nome_arquivo, pasta_certificados
        )
        caminho_db = salvar_dados_sqlite(
            dados_tela, df_final, val_k, val_u, calibracao_sucesso, pasta_dados
        )

        msg = f"Certificado salvo em `{caminho_pdf}` e dados gravados em `{caminho_db}`."
        if caminho_historico:
            msg += f" Versão anterior arquivada em `{caminho_historico}`."

        st.session_state["msg_persistencia"] = ("success", msg)
    except Exception as e:
        st.session_state["msg_persistencia"] = (
            "error",
            f"Erro ao salvar certificado/dados: {e}"
        )


def listar_anos_disponiveis(pasta_dados):
    """Lista os anos que possuem arquivo calibracoes_{ano}.db na pasta de dados."""
    anos = []
    if os.path.exists(pasta_dados):
        for arq in os.listdir(pasta_dados):
            if arq.startswith("calibracoes_") and arq.endswith(".db"):
                ano = arq.replace("calibracoes_", "").replace(".db", "")
                if ano.isdigit():
                    anos.append(ano)
    return sorted(anos, reverse=True)


def repair_xlsx_stream(in_bytes):
    """Corrige problemas de namespace em arquivos Excel gerados por máquinas."""
    in_buf = io.BytesIO(in_bytes)
    out_buf = io.BytesIO()
    with zipfile.ZipFile(in_buf, "r") as z_in, zipfile.ZipFile(out_buf, "w") as z_out:
        for item in z_in.infolist():
            data = z_in.read(item.filename)
            if item.filename.endswith(".xml"):
                data_str = data.decode("utf-8")
                data_str = data_str.replace(STRICT_NS, TRANS_NS)
                data = data_str.encode("utf-8")
            z_out.writestr(item, data)
    out_buf.seek(0)
    return out_buf


# Dicionários de opções para os menus dropdown
OPCOES_CLIENTE = ["Metrologia", "Usinagem", "Pintura", "Compressores", "Lab. Fundição", "Q. Fundição", "Fusão", "Macharia", "Modelação", "Moldagem", "Gestão Ambiental", "Acabamento", "Manutenção", "SHMT", "IRA", "U. Compressores", "Eng. Processo", "Fabrica II", "Poka Yoke", "Logística", "Lab. China", "Lab. Sujidade", "Presset/Ferrament.", "Usinagem II", "Comp. Seg. do Trab. e Manut.", "Controle Dim. I", "Controle Dim. II", "Pintura Usi. II", "Schulz RVA", "Fund.", "Unid. 3"]
OPCOES_METROLOGISTA = ["José Alcides Rosa Jr", "Cleocir de Souza"]
OPCOES_MMC = ["500.008", "500.009", "500.011", "500.012", "500.015", "500.016", "500.018", "500.022"]
OPCOES_FORNECEDOR = ["CDM", "Usifer", "Torque", "Valente"]


def init_session_state():
    chaves = [
        'Teste', 'Dispositivo de Controle', 'Data', 'Ordem', 'Temperatura',
        'Desenho', 'Hora', 'Numero peça', 'Cliente', 'Metrologista', 'MMC',
        'Fornecedor', 'df_final', 'calibracao_sucesso', 'val_k', 'val_u',
        'val_veff', 'processado', 'msg_persistencia'
    ]
    for chave in chaves:
        if chave not in st.session_state:
            st.session_state[chave] = "" if chave not in [
                'df_final', 'calibracao_sucesso', 'val_k', 'val_u',
                'val_veff', 'processado', 'msg_persistencia'
            ] else None
    if not st.session_state['Cliente']: st.session_state['Cliente'] = "Usinagem"
    if not st.session_state['MMC']: st.session_state['MMC'] = "500.011"


aplicar_estilo_customizado()
init_session_state()
pasta_certificados_atual, pasta_dados_atual = carregar_configuracao()


def calcular_incerteza(df_final):
    """
    Calcula o fator de abrangência (k), a incerteza expandida (U em mm)
    e os graus de liberdade efetivos (veff) conforme a planilha da Schulz.
    """
    comprimento_avaliado = df_final['Nominal'].max() if not df_final['Nominal'].isna().all() else 0.0
    n_medicoes = 2
    dif_temp_padrao_obj = 0.5
    var_max_temp_20 = 1.0

    max_desvio = df_final['Desvio padrão'].max() if not df_final['Desvio padrão'].isna().all() else 0.0
    s_um = max_desvio * 1000.0

    u_48 = (s_um / np.sqrt(n_medicoes)) * 1.0
    v_48 = (n_medicoes - 1) if n_medicoes > 1 else 1
    u_49 = (0.4 / 2.0) * 1.0
    v_49 = 1000
    u_50 = (0.6 / 2.0) * 1.0
    v_50 = 1000
    u_51 = (0.9 / np.sqrt(3)) * 1.0
    v_51 = 1000
    u_52 = (0.1 / (2.0 * np.sqrt(3))) * 1.0
    v_52 = 1000
    u_53 = (-0.8 / np.sqrt(3)) * 1.0
    v_53 = 1000

    b54 = var_max_temp_20 * 0.000002
    c54 = np.sqrt(3) * np.sqrt(6)
    f54 = comprimento_avaliado * 1000.0
    u_54 = (b54 / c54) * f54
    v_54 = 1000

    b55 = dif_temp_padrao_obj
    c55 = np.sqrt(3)
    f55 = comprimento_avaliado * 11.5 * 0.001
    u_55 = (b55 / c55) * f55
    v_55 = 1000

    componentes = [
        (u_48, v_48), (u_49, v_49), (u_50, v_50), (u_51, v_51),
        (u_52, v_52), (u_53, v_53), (u_54, v_54), (u_55, v_55)
    ]

    uc = np.sqrt(sum(u**2 for u, _ in componentes))
    denominador_veff = sum((u**4) / v for u, v in componentes)
    veff = int((uc**4) / denominador_veff) if denominador_veff > 0 else 1000

    if veff < 2:      k = 13.97
    elif veff < 3:    k = 4.53
    elif veff < 4:    k = 3.31
    elif veff < 5:    k = 2.87
    elif veff < 6:    k = 2.65
    elif veff < 7:    k = 2.52
    elif veff < 8:    k = 2.43
    elif veff < 10:   k = 2.37
    elif veff < 12:   k = 2.28
    elif veff < 14:   k = 2.23
    elif veff < 16:   k = 2.20
    elif veff < 18:   k = 2.17
    elif veff < 20:   k = 2.15
    elif veff < 25:   k = 2.13
    elif veff < 30:   k = 2.11
    elif veff < 35:   k = 2.09
    elif veff < 40:   k = 2.07
    elif veff < 50:   k = 2.06
    elif veff < 60:   k = 2.05
    elif veff < 80:   k = 2.04
    elif veff < 100:  k = 2.03
    elif veff == 100: k = 2.02
    else:             k = 2.00

    u_mm = (uc * k) / 1000.0
    return k, u_mm, veff


def processar_planilha(uploaded_file):
    try:
        stream_corrigido = repair_xlsx_stream(uploaded_file.getvalue())
        df_topo = pd.read_excel(stream_corrigido, sheet_name="Report", nrows=12, header=None, engine="openpyxl")

        st.session_state['Dispositivo de Controle'] = str(df_topo.iloc[4, 1]) if pd.notna(df_topo.iloc[4, 1]) else ""
        st.session_state['Data'] = str(df_topo.iloc[4, 3]).split()[0] if pd.notna(df_topo.iloc[4, 3]) else ""
        st.session_state['Ordem'] = str(df_topo.iloc[4, 5]) if pd.notna(df_topo.iloc[4, 5]) else ""
        st.session_state['Desenho'] = str(df_topo.iloc[7, 1]) if pd.notna(df_topo.iloc[7, 1]) else ""
        st.session_state['Hora'] = str(df_topo.iloc[7, 3]) if pd.notna(df_topo.iloc[7, 3]) else ""
        st.session_state['Numero peça'] = str(df_topo.iloc[7, 5]) if pd.notna(df_topo.iloc[7, 5]) else ""

        metrologista = str(df_topo.iloc[10, 1]) if pd.notna(df_topo.iloc[10, 1]) else ""
        if metrologista in OPCOES_METROLOGISTA: st.session_state['Metrologista'] = metrologista

        mmc = str(df_topo.iloc[10, 3]) if pd.notna(df_topo.iloc[10, 3]) else ""
        if mmc in OPCOES_MMC: st.session_state['MMC'] = mmc

        fornecedor = str(df_topo.iloc[10, 5]) if pd.notna(df_topo.iloc[10, 5]) else ""
        if fornecedor in OPCOES_FORNECEDOR: st.session_state['Fornecedor'] = fornecedor

        stream_corrigido.seek(0)
        df_report = pd.read_excel(stream_corrigido, sheet_name="Report", header=12, engine="openpyxl")
        df_report = df_report.loc[:, ~df_report.columns.str.contains('^Unnamed')]
        df_report = df_report.dropna(subset=["Characteristic"])

        colunas_num = ["Actual", "Nominal", "Upper Tol", "Lower Tol"]
        for col in colunas_num:
            if col in df_report.columns:
                df_report[col] = pd.to_numeric(df_report[col], errors='coerce')

        df_report['Medida_Num'] = df_report.groupby('Characteristic').cumcount() + 1
        df_medidas = df_report.pivot(index='Characteristic', columns='Medida_Num', values='Actual').reset_index()

        if 1 in df_medidas.columns: df_medidas.rename(columns={1: 'Medida 1'}, inplace=True)
        else: df_medidas['Medida 1'] = np.nan
        if 2 in df_medidas.columns: df_medidas.rename(columns={2: 'Medida 2'}, inplace=True)
        else: df_medidas['Medida 2'] = np.nan

        df_limites = df_report.drop_duplicates(subset=['Characteristic'])[['Characteristic', 'Nominal', 'Upper Tol', 'Lower Tol']]
        df_final = pd.merge(df_limites, df_medidas, on='Characteristic')
        df_final['Média'] = df_final[['Medida 1', 'Medida 2']].mean(axis=1)
        df_final['Desvio padrão'] = df_final[['Medida 1', 'Medida 2']].std(axis=1, ddof=1)

        calibracao_sucesso = True
        status_aprovado = []

        for _, row in df_final.iterrows():
            nominal, upper_tol, lower_tol, media = row['Nominal'], row['Upper Tol'], row['Lower Tol'], row['Média']
            aprovado = True

            if pd.notnull(nominal) and pd.notnull(media) and pd.notnull(upper_tol) and pd.notnull(lower_tol):
                limite_superior = nominal + upper_tol
                limite_inferior = nominal - lower_tol if lower_tol > 0 else nominal + lower_tol
                if not (limite_inferior <= media <= limite_superior):
                    aprovado = False
                    calibracao_sucesso = False

            status_aprovado.append(aprovado)

        df_final['Aprovado'] = status_aprovado

        val_k, val_u, val_veff = calcular_incerteza(df_final)
        st.session_state['val_k'] = val_k
        st.session_state['val_u'] = val_u
        st.session_state['val_veff'] = val_veff
        st.session_state['df_final'] = df_final
        st.session_state['calibracao_sucesso'] = calibracao_sucesso
        st.session_state['processado'] = True

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar a planilha: {str(e)}")


def gerar_pdf(dados_tela, df_final, val_k, val_u, val_veff=1000, calibracao_sucesso=True):
    dispositivo = dados_tela.get("Dispositivo de Controle", "N_A")
    ano, data_formatada = extrair_ano_e_data(dados_tela.get("Data", "N_A"))

    cliente = dados_tela.get("Cliente", "")
    metrologista = dados_tela.get("Metrologista", "")
    fornecedor = dados_tela.get("Fornecedor", "Não consta")
    mmc = dados_tela.get("MMC", "")

    ordem = str(dados_tela.get("Ordem", "00000")).strip()
    ordem_fmt = ordem.zfill(5) if ordem.isdigit() else ordem

    nome_arquivo = f"Certificado de Calibração {dispositivo}-{ordem_fmt}-{ano}.pdf"

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=1.5*cm,
        topMargin=1.5*cm,
        bottomMargin=1.5*cm
    )

    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(name='TitleStyle', parent=styles['Heading1'], alignment=1)
    style_normal = styles['Normal']
    style_bold = ParagraphStyle(name='BoldStyle', parent=styles['Normal'], fontName='Helvetica-Bold')

    elementos = []

    try:
        img = Image(os.path.join(BASE_DIR, 'Imagens', 'BarraSuperior.jpg'), width=18*cm, height=1.82*cm)
        img.hAlign = 'CENTER'
        elementos.append(img)
        elementos.append(Spacer(1, 12))
    except Exception:
        pass

    elementos.append(Paragraph(f"<b>Certificado de Calibração nº {ordem_fmt}/{ano}</b>   Data: {data_formatada}", style_title))
    elementos.append(Paragraph(f"<b>Solicitante:</b> {cliente}", style_normal))
    elementos.append(Spacer(1, 12))

    elementos.append(Paragraph("<b>1. Descrição do instrumento calibrado</b>", style_bold))
    elementos.append(Paragraph("Denominação: Dispositivo de controle", style_normal))
    elementos.append(Paragraph(f"Código: {dispositivo}", style_normal))
    elementos.append(Paragraph(f"Fabricante: {fornecedor}", style_normal))
    elementos.append(Spacer(1, 12))

    elementos.append(Paragraph("<b>2. Padrão utilizado</b>", style_bold))
    elementos.append(Paragraph(f"Máquina de medição por coordenadas {mmc} - Carl Zeiss", style_normal))
    elementos.append(Paragraph("Certificado de Calibração n° 241120502390 - Carl Zeiss; de 20/11/2024 - Próxima Calibração: 31/08/2026", style_normal))
    elementos.append(Spacer(1, 12))

    elementos.append(Paragraph("<b>3. Procedimento de calibração</b>", style_bold))
    elementos.append(Paragraph("Nr. do procedimento / revisão: Calibração de dispositivo de controle - ITL-0026 (revisão 2).", style_normal))
    elementos.append(Paragraph("O instrumento foi medido na máquina de medição por coordenadas através de elementos geométricos definidos, sendo o resultado apresentado a média de duas medições realizadas.", style_normal))
    elementos.append(Paragraph("Os resultados estão referenciados a temperatura de 20° Celsius.", style_normal))
    elementos.append(Spacer(1, 12))

    elementos.append(Paragraph("<b>4. Resultados (mm)</b>", style_bold))
    elementos.append(Spacer(1, 6))

    dados_tabela = [["Parâmetro avaliado", "Valor nominal", "Tol. +", "Tol. -", "Valor obtido", "Erro"]]

    for _, row in df_final.iterrows():
        parametro = row['Characteristic']
        nominal = row['Nominal']
        media_obtida = row['Média']

        try:
            erro_calculado = abs(float(media_obtida) - float(nominal))
            erro = f"{erro_calculado:.3f}".replace('.', ',')
        except (ValueError, TypeError):
            erro = "-"

        dados_tabela.append([
            str(parametro),
            f"{nominal:.3f}".replace('.', ',') if pd.notna(nominal) else "-",
            f"{row['Upper Tol']:.3f}".replace('.', ',') if pd.notna(row['Upper Tol']) else "-",
            f"{row['Lower Tol']:.3f}".replace('.', ',') if pd.notna(row['Lower Tol']) else "-",
            f"{media_obtida:.3f}".replace('.', ',') if pd.notna(media_obtida) else "-",
            erro
        ])

    tabela = Table(dados_tabela, colWidths=[5.5*cm, 2.5*cm, 2*cm, 2*cm, 2.5*cm, 2*cm])
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    elementos.append(tabela)
    elementos.append(Spacer(1, 12))

    val_u_fmt = f"{val_u:.4f}".replace('.', ',')
    val_k_fmt = f"{val_k:.2f}".replace('.', ',')
    elementos.append(Paragraph(f"<b>5. Incerteza de medição:</b> {val_u_fmt} mm, k = {val_k_fmt}", style_bold))
    elementos.append(Spacer(1, 12))

    elementos.append(Paragraph("<b>6. Observações</b>", style_bold))
    elementos.append(Paragraph("*Os parâmetros avaliados foram obtidos da Ficha de Auto Controle do produto que o dispositivo controla.", style_normal))
    elementos.append(Paragraph(f"Técnico executante: {metrologista}", style_normal))
    elementos.append(Spacer(1, 24))

    status_str = "APROVADO" if calibracao_sucesso else "REPROVADO"
    codigo_unico = f"{dispositivo}-{ordem_fmt}-{ano}"
    params_qr = urlencode({
        "id": codigo_unico,
        "dispositivo": dispositivo,
        "data": data_formatada,
        "status": status_str
    })
    url_validacao = f"https://schulz-metrologia.validacao/certificado?{params_qr}"

    qr_widget = qr.QrCodeWidget(url_validacao)
    bounds = qr_widget.getBounds()
    w = bounds[2] - bounds[0]
    h = bounds[3] - bounds[1]
    tamanho_qr = 3.0 * cm
    desenho_qr = Drawing(
        tamanho_qr, tamanho_qr,
        transform=[tamanho_qr / w, 0, 0, tamanho_qr / h, 0, 0]
    )
    desenho_qr.add(qr_widget)

    style_qr_legenda = ParagraphStyle(
        name='QRLegenda',
        parent=styles['Normal'],
        fontSize=7,
        alignment=1,
        leading=9
    )

    bloco_assinatura = [
        Paragraph("_______________________________", style_normal),
        Paragraph("João Carlos Gomes de Oliveira", style_normal),
        Paragraph("Signatário Autorizado", style_normal),
    ]

    bloco_qr = [
        desenho_qr,
        Paragraph(f"<b>Validação Mobile</b><br/>{codigo_unico}", style_qr_legenda)
    ]

    tabela_rodape = Table([[bloco_assinatura, bloco_qr]], colWidths=[14.0 * cm, 4.0 * cm])
    tabela_rodape.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))

    elementos.append(tabela_rodape)

    # ==========================================
    # NOTAS E RODAPÉ INSTITUCIONAL (PÓS-ASSINATURA)
    # ==========================================
    elementos.append(Spacer(1, 10))

    style_notas = ParagraphStyle(
        name='NotasStyle',
        parent=styles['Normal'],
        fontSize=7.5,
        leading=9.5,
        alignment=0
    )
    style_rodape_inst = ParagraphStyle(
        name='RodapeInstStyle',
        parent=styles['Normal'],
        fontSize=7.5,
        leading=10,
        alignment=1
    )

    veff_fmt = str(int(val_veff)) if val_veff is not None else "1000"

    elementos.append(HRFlowable(width="100%", thickness=0.8, color=colors.black, spaceBefore=2, spaceAfter=4))

    elementos.append(Paragraph("<b>Notas:</b>", style_notas))
    elementos.append(Paragraph(
        f"* A incerteza expandida de medição relatada é declarada como incerteza padrão da medição multiplicada pelo "
        f"fator de abrangência <b>k = {val_k_fmt}</b>, o qual para uma distribuição t com <b>veff = {veff_fmt}</b> "
        f"graus de liberdade efetivos corresponde a uma probabilidade de abrangência de aproximadamente 95% "
        f"(k e veff apresentados acima). A incerteza padrão de medição foi determinada de acordo com a publicação EA-4/02.",
        style_notas
    ))
    elementos.append(Paragraph(
        "* Os resultados deste certificado referem-se exclusivamente ao instrumento submetido a calibração, "
        "nas condições especificadas, não sendo extensivo a quaisquer lotes.",
        style_notas
    ))
    elementos.append(Paragraph("* Não é permitida a reprodução parcial deste certificado.", style_notas))
    elementos.append(Paragraph("* Finalização do relatório.", style_notas))

    elementos.append(HRFlowable(width="100%", thickness=0.8, color=colors.black, spaceBefore=4, spaceAfter=2))
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=colors.black, spaceBefore=0, spaceAfter=4))

    elementos.append(Paragraph("Folha 01/01", style_rodape_inst))
    elementos.append(Paragraph(
        "SCHULZ Metrologia - Rua Dona Francisca, 6901 - Distrito Industrial • 89219-600 • Fone (47) 3451-6342 • Joinville • SC",
        style_rodape_inst
    ))
    elementos.append(Paragraph(
        "C.N.P.J. 84.684.182/0001-57 • Insc. Est. 250.135.558 • e-mail metrologia@schulz.com.br • www.schulz.com.br",
        style_rodape_inst
    ))

    elementos.append(Spacer(1, 6))
    elementos.append(HRFlowable(width="100%", thickness=6, color=colors.HexColor("#002b5c"), spaceBefore=2, spaceAfter=0))

    doc.build(elementos)
    pdf_buffer.seek(0)
    return pdf_buffer, nome_arquivo


# ==========================================
# BARRA LATERAL: IDENTIDADE, EXPORTAÇÃO E DIRETÓRIOS
# ==========================================
with st.sidebar:
    caminho_barra = os.path.join(BASE_DIR, "Imagens", "BarraSuperior.jpg")
    if os.path.exists(caminho_barra):
        st.image(caminho_barra, use_container_width=True)
        st.markdown("---")

    st.header("🗄️ Banco de Dados & Relatórios")
    anos_disponiveis = listar_anos_disponiveis(pasta_dados_atual)

    if anos_disponiveis:
        ano_sel = st.selectbox("Selecione o ano para exportar:", anos_disponiveis)
        if st.button("📊 Exportar Relatório Excel (.xlsx)", use_container_width=True):
            try:
                caminho_excel, nome_excel, excel_bytes = exportar_sqlite_para_excel(
                    ano_sel, pasta_dados_atual
                )
                st.success(f"Exportado em:\n`{caminho_excel}`")
                st.download_button(
                    label=f"⬇️ Baixar {nome_excel}",
                    data=excel_bytes,
                    file_name=nome_excel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except PermissionError:
                st.error(f"Feche o arquivo `Calibrações-{ano_sel}.xlsx` no Excel antes de exportar novamente.")
            except Exception as e:
                st.error(f"Erro ao exportar Excel: {e}")
    else:
        st.info("Nenhuma calibração registrada ainda. O banco SQLite será criado ao salvar o primeiro certificado.")

    st.markdown("---")
    st.caption("📁 **Diretórios Ativos (`config.txt`)**")
    st.caption(f"**Certificados:** `{pasta_certificados_atual}`")
    st.caption(f"**Dados (SQLite/Excel):** `{pasta_dados_atual}`")


# ==========================================
# BANNER PRINCIPAL COM IMAGEM DA MMC
# ==========================================
img_mmc_b64 = carregar_imagem_base64(os.path.join("Imagens", "MMC7.JPG"))

if img_mmc_b64:
    st.markdown(
        f"""
        <div class="hero-banner">
            <div style="display: flex; align-items: center; justify-content: space-between; gap: 24px; flex-wrap: wrap;">
                <div style="flex: 2; min-width: 280px;">
                    <div class="hero-title">Automação de Certificados CMM — Schulz S.A.</div>
                    <div class="hero-subtitle">
                        Ingestão de relatórios da Máquina de Medição por Coordenadas, cálculo automático de incerteza expandida (U),
                        fator de abrangência (k) e emissão padronizada de certificados com rastreabilidade.
                    </div>
                    <span class="hero-badge">📐 ITL-0026 rev. 2</span>
                    <span class="hero-badge">🗄️ Persistência SQLite + Excel</span>
                    <span class="hero-badge">📱 Validação QR Code</span>
                </div>
                <div style="flex: 1; min-width: 240px; max-width: 340px;">
                    <div class="mmc-img-wrapper">
                        <img src="data:image/jpeg;base64,{img_mmc_b64}" alt="MMC Carl Zeiss Contura" />
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        """
        <div class="hero-banner">
            <div class="hero-title">Automação de Certificados CMM — Schulz S.A.</div>
            <div class="hero-subtitle">
                Ingestão de relatórios da Máquina de Medição por Coordenadas, cálculo automático de incerteza expandida (U)
                e emissão padronizada de certificados com rastreabilidade.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==========================================
# SEÇÃO 1: UPLOAD DA PLANILHA
# ==========================================
with st.container(border=True):
    st.markdown('<div class="section-title">1. Carregar Planilha de Medições</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Selecione ou arraste o relatório exportado pela CMM (.xlsx ou .xls)",
        type=["xlsx", "xls"],
        label_visibility="collapsed"
    )

if uploaded_file is None:
    if st.session_state.get('processado'):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        init_session_state()
        st.rerun()
else:
    if st.session_state.get('nome_arquivo_atual') != uploaded_file.name:
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        init_session_state()
        st.session_state['nome_arquivo_atual'] = uploaded_file.name
        processar_planilha(uploaded_file)
        st.rerun()


# ==========================================
# SEÇÃO 2: METADADOS DO CERTIFICADO
# ==========================================
with st.container(border=True):
    st.markdown('<div class="section-title">2. Metadados do Certificado</div>', unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        teste = st.text_input("Teste", value=st.session_state['Teste'])
        temp = st.text_input("Temperatura (°C)", value=st.session_state['Temperatura'])
        cliente = st.selectbox("Cliente", options=OPCOES_CLIENTE, index=OPCOES_CLIENTE.index(st.session_state['Cliente']) if st.session_state['Cliente'] in OPCOES_CLIENTE else 1)

    with col2:
        dispositivo = st.text_input("Dispositivo de Controle", value=st.session_state['Dispositivo de Controle'])
        desenho = st.text_input("Desenho", value=st.session_state['Desenho'])
        metro = st.selectbox("Metrologista", options=OPCOES_METROLOGISTA, index=OPCOES_METROLOGISTA.index(st.session_state['Metrologista']) if st.session_state['Metrologista'] in OPCOES_METROLOGISTA else 0)

    with col3:
        data_med = st.text_input("Data (dd/mm/aaaa)", value=st.session_state['Data'])
        hora = st.text_input("Hora (HH:MM)", value=st.session_state['Hora'])
        mmc_sel = st.selectbox("MMC", options=OPCOES_MMC, index=OPCOES_MMC.index(st.session_state['MMC']) if st.session_state['MMC'] in OPCOES_MMC else 2)

    with col4:
        ordem = st.text_input("Ordem", value=st.session_state['Ordem'])
        peca = st.text_input("Número peça", value=st.session_state['Numero peça'])
        fornecedor = st.selectbox("Fornecedor", options=OPCOES_FORNECEDOR, index=OPCOES_FORNECEDOR.index(st.session_state['Fornecedor']) if st.session_state['Fornecedor'] in OPCOES_FORNECEDOR else 0)


# ==========================================
# SEÇÃO 3: RESULTADOS, KPIS E EMISSÃO DO PDF
# ==========================================
if st.session_state.get('processado'):
    df_view = st.session_state['df_final'].copy()
    total_params = len(df_view)
    total_aprovados = int(df_view['Aprovado'].sum())
    status_aprovado = st.session_state['calibracao_sucesso']

    with st.container(border=True):
        st.markdown('<div class="section-title">3. Parâmetros Lidos e Consolidados</div>', unsafe_allow_html=True)

        def color_rows(row):
            color = '#dcfce7' if row['Aprovado'] else '#fee2e2'
            return [f'background-color: {color}; color: #0f172a' for _ in row]

        format_dict = {col: "{:.3f}" for col in ["Nominal", "Upper Tol", "Lower Tol", "Medida 1", "Medida 2", "Média", "Desvio padrão"]}
        st.dataframe(
            df_view.style.apply(color_rows, axis=1).format(format_dict, na_rep="-"),
            use_container_width=True,
            hide_index=True
        )

        st.markdown(
            """
            <div style="display: flex; gap: 24px; padding-top: 8px; padding-bottom: 4px;">
                <div style="display: flex; align-items: center;">
                    <div style="width: 14px; height: 14px; background-color: #dcfce7; border: 1px solid #86efac; border-radius: 3px; margin-right: 8px;"></div>
                    <span style="font-size: 13px; font-weight: 600; color: #166534;">DENTRO DA TOLERÂNCIA</span>
                </div>
                <div style="display: flex; align-items: center;">
                    <div style="width: 14px; height: 14px; background-color: #fee2e2; border: 1px solid #fca5a5; border-radius: 3px; margin-right: 8px;"></div>
                    <span style="font-size: 13px; font-weight: 600; color: #991b1b;">FORA DA TOLERÂNCIA</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    # Cards de KPIs
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    classe_status = "kpi-aprovado" if status_aprovado else "kpi-reprovado"
    texto_status = "✅ APROVADO" if status_aprovado else "❌ REPROVADO"

    with kpi1:
        st.markdown(
            f'<div class="kpi-card {classe_status}"><div class="kpi-label">Parecer Final</div><div class="kpi-value">{texto_status}</div></div>',
            unsafe_allow_html=True
        )
    with kpi2:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Parâmetros Conformes</div><div class="kpi-value">{total_aprovados} / {total_params}</div></div>',
            unsafe_allow_html=True
        )
    with kpi3:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Fator de Abrangência (k)</div><div class="kpi-value">{st.session_state["val_k"]:.2f}</div></div>',
            unsafe_allow_html=True
        )
    with kpi4:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Incerteza Expandida (U)</div><div class="kpi-value">{st.session_state["val_u"]:.4f} mm</div></div>',
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    dados_tela = {
        "Dispositivo de Controle": dispositivo,
        "Data": data_med,
        "Hora": hora,
        "Ordem": ordem,
        "Desenho": desenho,
        "Numero peça": peca,
        "Cliente": cliente,
        "Metrologista": metro,
        "MMC": mmc_sel,
        "Fornecedor": fornecedor,
        "Temperatura": temp
    }

    pdf_buffer, nome_arquivo = gerar_pdf(
        dados_tela,
        df_view,
        st.session_state['val_k'],
        st.session_state['val_u'],
        st.session_state['val_veff'],
        st.session_state['calibracao_sucesso']
    )

    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        st.download_button(
            label="📄 Salvar e Baixar Certificado em PDF",
            data=pdf_buffer.getvalue(),
            file_name=nome_arquivo,
            mime="application/pdf",
            use_container_width=True,
            on_click=salvar_certificado_e_dados,
            args=(
                pdf_buffer.getvalue(),
                nome_arquivo,
                dados_tela,
                df_view,
                st.session_state['val_k'],
                st.session_state['val_u'],
                st.session_state['calibracao_sucesso']
            )
        )

    if st.session_state.get("msg_persistencia"):
        tipo, texto = st.session_state["msg_persistencia"]
        if tipo == "success":
            st.success(texto)
        else:
            st.error(texto)