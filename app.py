import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
import shutil
import subprocess
import tempfile
import pandas as pd
import matplotlib.pyplot as plt
from streamlit_paste_button import paste_image_button
from PIL import Image
import platform
import time
import calendar
import json
from pathlib import Path
import zipfile

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="RELATÓRIO ASSISTENCIAL MENSAL - NOVA CIDADE", layout="wide")

# --- CONSTANTES DO CONTRATO ---
META_DIARIA_CONTRATO = 250

# --- CUSTOM CSS ---
st.markdown("""
    <style>
    .main { background-color: #f0f2f5; }
    .dashboard-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
        border-left: 5px solid #28a745;
    }
    div.stButton > button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
        font-weight: bold !important;
        height: 3em !important;
        border-radius: 8px !important;
    }
    div.stButton > button[key*="del_"] {
        border: 1px solid #dc3545 !important;
        color: #dc3545 !important;
        background-color: transparent !important;
        font-size: 0.8em !important;
        height: 2em !important;
    }
    div.stButton > button[key*="btn_"] {
        background-color: #ffffff !important;
        color: #374151 !important;
        border: 1px solid #d1d5db !important;
        border-radius: 6px !important;
        width: 100% !important;
        height: 2.5em !important;
    }
    .upload-label { font-weight: bold; color: #1f2937; margin-bottom: 8px; display: block; }
    </style>
    """, unsafe_allow_html=True)

# --- DICIONÁRIO DE DIMENSÕES ---
DIMENSOES_CAMPOS = {
    "IMAGEM_PRINT_ATENDIMENTO": 165, "PRINT_CLASSIFICACAO": 160,
    "IMAGEM_DOCUMENTO_RAIO_X": 150, "TABELA_TRANSFERENCIA": 90,
    "GRAFICO_TRANSFERENCIA": 150, "TABELA_OBITO": 180, 
    "TABELA_CCIH": 175, "IMAGEM_NEP": 160,
    "IMAGEM_TREINAMENTO_INTERNO": 160, "IMAGEM_MELHORIAS": 160,
    "GRAFICO_OUVIDORIA": 155, "PDF_OUVIDORIA_INTERNA": 165,
    "TABELA_QUALITATIVA_IMG": 190
}

# --- CONFIGURAÇÃO DE PERSISTÊNCIA ---
BASE_RELATORIOS_DIR = Path("relatorios_salvos_novacidade")
BASE_RELATORIOS_DIR.mkdir(exist_ok=True)

FORM_KEYS = [
    "sel_mes", "sel_ano", "in_total", "in_rx", "in_mc", "in_mp",
    "in_oc", "in_op", "in_ccih", "in_oi", "in_oe", "in_taxa",
    "in_tt", "in_to", "in_to_menor", "in_to_maior"
]

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}

if 'relatorio_atual' not in st.session_state:
    st.session_state.relatorio_atual = ""

# --- FUNÇÕES DE PERSISTÊNCIA ---

def _normalizar_nome(nome):
    return "".join([c if c.isalnum() else "_" for c in nome])

def listar_relatorios_salvos():
    return sorted([p.name for p in BASE_RELATORIOS_DIR.iterdir() if p.is_dir()])

def salvar_relatorio(nome):
    if not nome: return
    nome_norm = _normalizar_nome(nome)
    pasta = BASE_RELATORIOS_DIR / nome_norm
    pasta.mkdir(parents=True, exist_ok=True)
    pasta_evid = pasta / "evidencias"
    pasta_evid.mkdir(exist_ok=True)

    evid_meta = {}
    for m, itens in st.session_state.dados_sessao.items():
        evid_meta[m] = []
        for i, item in enumerate(itens):
            ext = ".png"
            fname = f"{m}_{i}{ext}"
            caminho_dest = pasta_evid / fname
            conteudo = item["content"]
            if isinstance(conteudo, Image.Image):
                conteudo.save(caminho_dest, format="PNG")
            else:
                if hasattr(conteudo, "seek"): conteudo.seek(0) # Volta ao início
                if hasattr(conteudo, "getvalue"): data = conteudo.getvalue()
                elif hasattr(conteudo, "read"): 
                    data = conteudo.read()
                else: data = conteudo
                with open(caminho_dest, "wb") as f: f.write(data)
            evid_meta[m].append({"name": item["name"], "file": f"evidencias/{fname}", "type": item["type"]})

    estado = {"form_state": {k: st.session_state.get(k) for k in FORM_KEYS}, "evidencias": evid_meta}
    with open(pasta / "estado.json", "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)
    st.session_state.relatorio_atual = nome_norm
    st.success(f"Relatório '{nome}' salvo com sucesso!")

def carregar_relatorio(nome_pasta):
    pasta = BASE_RELATORIOS_DIR / nome_pasta
    estado_path = pasta / "estado.json"
    if not estado_path.exists(): return
    with open(estado_path, "r", encoding="utf-8") as f:
        estado = json.load(f)
    for k, v in estado.get("form_state", {}).items():
        st.session_state[k] = v
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
    for m, lista in estado.get("evidencias", {}).items():
        for meta in lista:
            p = pasta / meta["file"]
            if p.exists():
                with open(p, "rb") as f:
                    bio = io.BytesIO(f.read())
                    bio.name = meta["name"]
                    st.session_state.dados_sessao[m].append({"name": meta["name"], "content": bio, "type": meta["type"]})
    st.session_state.relatorio_atual = nome_pasta
    st.success(f"Relatório carregado!")

# --- FUNÇÕES DE EXPORTAR E IMPORTAR (ZIP) ---

def gerar_backup_zip():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        evid_meta = {}
        for marcador, itens in st.session_state.dados_sessao.items():
            evid_meta[marcador] = []
            for i, item in enumerate(itens):
                conteudo = item["content"]
                file_bytes = b""
                if isinstance(conteudo, Image.Image):
                    img_buf = io.BytesIO()
                    conteudo.save(img_buf, format="PNG")
                    file_bytes = img_buf.getvalue()
                else:
                    if hasattr(conteudo, "seek"): conteudo.seek(0) # Reset do ponteiro
                    if hasattr(conteudo, "getvalue"): file_bytes = conteudo.getvalue()
                    elif hasattr(conteudo, "read"): 
                        file_bytes = conteudo.read()
                    else: file_bytes = conteudo
                
                nome_interno = f"evidencias/{marcador}_{i}.png"
                zf.writestr(nome_interno, file_bytes)
                evid_meta[marcador].append({"name": item["name"], "file": nome_interno, "type": item["type"]})
        
        estado = {"form_state": {k: st.session_state.get(k) for k in FORM_KEYS}, "evidencias": evid_meta}
        zf.writestr("estado.json", json.dumps(estado, ensure_ascii=False, indent=2))
    buf.seek(0)
    return buf

def processar_upload_backup(uploaded_zip):
    try:
        with zipfile.ZipFile(uploaded_zip, "r") as zf:
            estado_str = zf.read("estado.json").decode("utf-8")
            estado = json.loads(estado_str)
            for k, v in estado.get("form_state", {}).items():
                st.session_state[k] = v
            st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
            for marcador, lista in estado.get("evidencias", {}).items():
                for meta in lista:
                    try:
                        file_bytes = zf.read(meta["file"])
                        bio = io.BytesIO(file_bytes)
                        bio.name = meta["name"]
                        st.session_state.dados_sessao[marcador].append({"name": meta["name"], "content": bio, "type": meta["type"]})
                    except Exception: pass
        st.success("✅ Backup importado com sucesso!")
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro de backup: {e}")

# --- FUNÇÕES CORE ---

def converter_para_pdf(docx_path, output_dir):
    comando = 'libreoffice'
    if platform.system() == "Windows":
        caminhos = ['libreoffice', r'C:\Program Files\LibreOffice\program\soffice.exe', r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']
        for p in caminhos:
            try:
                subprocess.run([p, '--version'], capture_output=True, check=True)
                comando = p
                break
            except: continue
    subprocess.run([comando, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path], check=True)

def processar_item_lista(doc_template, item, marcador):
    largura = DIMENSOES_CAMPOS.get(marcador, 165)
    try:
        if hasattr(item, 'seek'): item.seek(0) # Garante que está no começo
        if isinstance(item, Image.Image):
            img_buf = io.BytesIO()
            item.save(img_buf, format='PNG')
            img_buf.seek(0)
            return [InlineImage(doc_template, img_buf, width=Mm(largura))]
        if isinstance(item, bytes):
            return [InlineImage(doc_template, io.BytesIO(item), width=Mm(largura))]
        if hasattr(item, 'seek'): item.seek(0)
        ext = getattr(item, 'name', '').lower()
        if ext.endswith(".pdf"):
            pdf = fitz.open(stream=item.read(), filetype="pdf")
            imgs = []
            for pg in pdf:
                pix = pg.get_pixmap(matrix=fitz.Matrix(2, 2))
                imgs.append(InlineImage(doc_template, io.BytesIO(pix.tobytes()), width=Mm(largura)))
            pdf.close()
            return imgs
        return [InlineImage(doc_template, item, width=Mm(largura))]
    except Exception: return []

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3208/3208726.png", width=100)
    st.title("Painel de Controle")
    st.markdown("---")
    total_anexos = sum(len(v) for v in st.session_state.dados_sessao.values())
    st.metric("Total de Anexos", total_anexos)
    if st.button("🗑 Limpar Todos os Dados"):
        st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
        st.rerun()

# --- UI PRINCIPAL ---
st.title("Automação de Relatórios - UPA Nova Cidade")

# --- BACKUP DE SEGURANÇA ---
with st.container(border=True):
    st.markdown("#### ☁️ Backup de Segurança (Exportar / Importar)")
    st.caption("Utilize esta opção para não perder os seus dados caso o servidor reinicie.")
    col_up, col_down = st.columns(2)
    with col_up:
        zip_upload = st.file_uploader("📥 Retomar Relatório (Carregar .zip)", type=["zip"], key="upload_backup")
        if zip_upload:
            if st.button("Restaurar Dados do ZIP", key="btn_restore", use_container_width=True):
                processar_upload_backup(zip_upload)
                time.sleep(1)
                st.rerun()
    with col_down:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        zip_buffer = gerar_backup_zip()
        nome_backup = f"Backup_Relatorio_{st.session_state.get('sel_mes', 'Atual')}.zip"
        st.download_button(
            label="📤 Guardar Progresso (Baixar .zip)",
            data=zip_buffer,
            file_name=nome_backup,
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

st.caption("Versão 0.7.12")

t_manual, t_evidencia = st.tabs(["Dados", "Evidências"])

with t_manual:
    st.markdown("### Configuração do Período e Metas")
    c1, c2, c3 = st.columns(3)
    meses_pt = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    with c1: mes_selecionado = st.selectbox("Mês de Referência", meses_pt, key="sel_mes")
    with c2: ano_selecionado = st.selectbox("Ano", [2025, 2026, 2027, 2028], index=2, key="sel_ano")
    with c3: st.text_input("Total de Atendimentos", key="in_total")

    mes_num = meses_pt.index(mes_selecionado) + 1
    dias_no_mes = calendar.monthrange(ano_selecionado, mes_num)[1]
    meta_calculada = dias_no_mes * META_DIARIA_CONTRATO
    meta_min = int(meta_calculada * 0.75)
    meta_max = int(meta_calculada * 1.25)

    c4, c5, c6 = st.columns(3)
    with c4: st.text_input("Meta do Mês (Calculada)", value=str(meta_calculada), disabled=True)
    with c5: st.text_input("Meta -25% (Calculada)", value=str(meta_min), disabled=True)
    with c6: st.text_input("Meta +25% (Calculada)", value=str(meta_max), disabled=True)

    st.markdown("---")
    st.markdown("### Dados Assistenciais")
    c7, c8, c9 = st.columns(3)
    with c7: st.text_input("Total Raio-X", key="in_rx")
    with c8: st.text_input("Médicos Clínicos", key="in_mc")
    with c9: st.text_input("Médicos Pediatras", key="in_mp")
    c10, c11, c12 = st.columns(3)
    with c10: st.text_input("Odonto Clínico", key="in_oc")
    with c11: st.text_input("Odonto Ped", key="in_op")
    with c12: st.text_input("Pacientes CCIH", key="in_ccih")
    c13, c14, c15 = st.columns(3)
    with c13: st.text_input("Ouvidoria Interna", key="in_oi")
    with c14: st.text_input("Ouvidoria Externa", key="in_oe")
    with c15: st.text_input("Taxa de Transferência (%)", key="in_taxa")
    c16, c17, c18 = st.columns(3)
    with c16: st.number_input("Total de Transferências", step=1, key="in_tt")
    with c17: st.number_input("Total de Óbitos", key="in_to", step=1)
    with c18: st.number_input("Óbito < 24h", key="in_to_menor", step=1)
    c19, c20, c21 = st.columns(3)
    with c19: st.number_input("Óbito > 24h", key="in_to_maior", step=1)

with t_evidencia:
    labels = {
        "IMAGEM_PRINT_ATENDIMENTO": "Prints Atendimento", "PRINT_CLASSIFICACAO": "Classificação de Risco", 
        "IMAGEM_DOCUMENTO_RAIO_X": "Doc. Raio-X", "TABELA_TRANSFERENCIA": "Tabela Transferência", 
        "GRAFICO_TRANSFERENCIA": "Gráfico Transferência", "TABELA_OBITO": "Tab. Óbito", 
        "TABELA_CCIH": "Tabela CCIH", "TABELA_QUALITATIVA_IMG": "Tab. Qualitativa",
        "IMAGEM_NEP": "Imagens NEP", "IMAGEM_TREINAMENTO_INTERNO": "Treinamento Interno", 
        "IMAGEM_MELHORIAS": "Melhorias", "GRAFICO_OUVIDORIA": "Gráfico Ouvidoria", "PDF_OUVIDORIA_INTERNA": "Relatório Ouvidoria"
    }
    blocos = [
        ["IMAGEM_PRINT_ATENDIMENTO", "PRINT_CLASSIFICACAO", "IMAGEM_DOCUMENTO_RAIO_X"],
        ["TABELA_TRANSFERENCIA", "GRAFICO_TRANSFERENCIA"],
        ["TABELA_OBITO", "TABELA_CCIH", "TABELA_QUALITATIVA_IMG"],
        ["IMAGEM_NEP", "IMAGEM_TREINAMENTO_INTERNO", "IMAGEM_MELHORIAS", "GRAFICO_OUVIDORIA", "PDF_OUVIDORIA_INTERNA"]
    ]
    for b_idx, lista_m in enumerate(blocos):
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        col_esq, col_dir = st.columns(2)
        for idx, m in enumerate(lista_m):
            target = col_esq if idx % 2 == 0 else col_dir
            with target:
                st.markdown(f"<span class='upload-label'>{labels.get(m, m)}</span>", unsafe_allow_html=True)
                ca, cb = st.columns([1, 1])
                with ca:
                    key_p = f"p_{m}_{len(st.session_state.dados_sessao[m])}"
                    pasted = paste_image_button(label="Colar Print", key=key_p)
                    if pasted and pasted.image_data:
                        st.session_state.dados_sessao[m].append({"name": f"Captura_{len(st.session_state.dados_sessao[m])+1}.png", "content": pasted.image_data, "type": "p"})
                        st.rerun()
                with cb:
                    f_up = st.file_uploader("Upload", type=['png', 'jpg', 'pdf', 'xlsx'], key=f"f_{m}_{b_idx}", label_visibility="collapsed")
                    if f_up:
                        arquivos_ja_anexados = [item['name'] for item in st.session_state.dados_sessao[m]]
                        if f_up.name not in arquivos_ja_anexados:
                            st.session_state.dados_sessao[m].append({"name": f_up.name, "content": f_up, "type": "f"})
                            st.rerun()
                if st.session_state.dados_sessao[m]:
                    for i_idx, item in enumerate(st.session_state.dados_sessao[m]):
                        with st.expander(f"{item['name']}", expanded=False):
                            if item['type'] == "p" or item['name'].lower().endswith(('.png', '.jpg', '.jpeg')):
                                # Reset do ponteiro antes de exibir a imagem
                                if hasattr(item['content'], 'seek'): item['content'].seek(0)
                                st.image(item['content'], use_container_width=True)
                            if st.button("Remover", key=f"del_{m}_{i_idx}_{b_idx}"):
                                st.session_state.dados_sessao[m].pop(i_idx)
                                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# --- GERAÇÃO FINAL ---
if st.button(" FINALIZAR E GERAR RELATÓRIO", type="primary"):
    try:
        with tempfile.TemporaryDirectory() as tmp:
            docx_p = os.path.join(tmp, "relatorio.docx")
            doc = DocxTemplate("template-upa-nova-cidade.docx")
            dados_finais = {
                "SISTEMA_MES_REFERENCIA": f"{st.session_state.get('sel_mes', 'Janeiro')}/{st.session_state.get('sel_ano', 2026)}",
                "ANALISTA_TOTAL_ATENDIMENTOS": st.session_state.get("in_total", ""),
                "TOTAL_RAIO_X": st.session_state.get("in_rx", ""),
                "ANALISTA_META_MES": str(meta_calculada),
                "ANALISTA_META_MINUS_25": str(meta_min),
                "ANALISTA_META_PLUS_25": str(meta_max),
                "ANALISTA_MEDICO_CLINICO": st.session_state.get("in_mc", ""),
                "ANALISTA_MEDICO_PEDIATRA": st.session_state.get("in_mp", ""),
                "ANALISTA_ODONTO_CLINICO": st.session_state.get("in_oc", ""),
                "ANALISTA_ODONTO_PED": st.session_state.get("in_op", ""),
                "TOTAL_PACIENTES_CCIH": st.session_state.get("in_ccih", ""),
                "OUVIDORIA_INTERNA": st.session_state.get("in_oi", ""),
                "OUVIDORIA_EXTERNA": st.session_state.get("in_oe", ""),
                "SISTEMA_TOTAL_DE_TRANSFERENCIA": st.session_state.get("in_tt", 0),
                "SISTEMA_TAXA_DE_TRANSFERENCIA": st.session_state.get("in_taxa", ""),
                "ANALISTA_TOTAL_OBITO": st.session_state.get("in_to", 0),
                "ANALISTA_OBITO_MENOR": st.session_state.get("in_to_menor", 0),
                "ANALISTA_OBITO_MAIOR": st.session_state.get("in_to_maior", 0),
                "SISTEMA_TOTAL_MEDICOS": int(st.session_state.get("in_mc", 0) or 0) + int(st.session_state.get("in_mp", 0) or 0)
            }
            for m in DIMENSOES_CAMPOS.keys():
                imgs = []
                for item in st.session_state.dados_sessao.get(m, []):
                    res = processar_item_lista(doc, item['content'], m)
                    if res: imgs.extend(res)
                dados_finais[m] = imgs
            
            doc.render(dados_finais)
            doc.save(docx_p)
            st.success("✅ Relatório gerado!")
            c_down1, c_down2 = st.columns(2)
            with c_down1:
                with open(docx_p, "rb") as f_w:
                    st.download_button(label="Baixar WORD (.docx)", data=f_w.read(), file_name=f"RELATÓRIO_NOVA_CIDADE_{st.session_state.get('sel_mes')}.docx")
            with c_down2:
                try:
                    converter_para_pdf(docx_p, tmp)
                    pdf_p = os.path.join(tmp, "relatorio.pdf")
                    if os.path.exists(pdf_p):
                        with open(pdf_p, "rb") as f_p:
                            st.download_button(label="Baixar PDF", data=f_p.read(), file_name=f"RELATÓRIO_NOVA_CIDADE_{st.session_state.get('sel_mes')}.pdf")
                except: st.warning("Conversão PDF indisponível.")
    except Exception as e: st.error(f"Erro na geração: {e}")

st.caption("Desenvolvido por Leonardo Barcelos Martins")
