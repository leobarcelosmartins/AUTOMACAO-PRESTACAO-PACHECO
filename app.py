import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
import json
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
from pathlib import Path

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios - UPA Pacheco", layout="wide")

# --- CONFIGURAÇÃO DE PERSISTÊNCIA (v0.7.12) ---
BASE_RELATORIOS_DIR = Path("relatorios_salvos_pacheco")
BASE_RELATORIOS_DIR.mkdir(exist_ok=True)

# Chaves (keys) de todos os inputs que devem ser guardados
FORM_KEYS = [
    "sel_mes", "sel_ano", "in_total", "in_rx", "in_mc", "in_mp",
    "in_oc", "in_op", "in_ccih", "in_oi", "in_oe", "in_taxa",
    "in_tt", "in_to", "in_to_menor", "in_to_maior"
]

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
    .upload-label { font-weight: bold; color: #1f2937; margin-bottom: 8px; display: block; }
    </style>
    """, unsafe_allow_html=True)

# --- DICIONÁRIO DE DIMENSÕES ---
DIMENSOES_CAMPOS = {
    "IMAGEM_PRINT_ATENDIMENTO": 165, "PRINT_CLASSIFICACAO": 160,
    "IMAGEM_DOCUMENTO_RAIO_X": 165, "TABELA_TRANSFERENCIA": 120,
    "GRAFICO_TRANSFERENCIA": 60, "TABELA_OBITO": 180, 
    "TABELA_CCIH": 160, "IMAGEM_NEP": 160,
    "IMAGEM_MELHORIAS": 160, "GRAFICO_OUVIDORIA": 155, 
    "PDF_OUVIDORIA_INTERNA": 60, "TABELA_QUALITATIVA_IMG": 190
}

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
if 'relatorio_atual' not in st.session_state:
    st.session_state.relatorio_atual = ""

# --- FUNÇÕES DE PERSISTÊNCIA (v0.7.12) ---

def salvar_relatorio(nome):
    if not nome: return
    nome_norm = "".join([c if c.isalnum() else "_" for c in nome])
    pasta = BASE_RELATORIOS_DIR / nome_norm
    pasta.mkdir(parents=True, exist_ok=True)
    
    pasta_evid = pasta / "evidencias"
    pasta_evid.mkdir(exist_ok=True)
    
    evid_meta = {}
    for marcador, itens in st.session_state.dados_sessao.items():
        evid_meta[marcador] = []
        for i, item in enumerate(itens):
            ext = ".png"
            fname = f"{marcador}_{i}{ext}"
            caminho_dest = pasta_evid / fname
            conteudo = item["content"]
            
            if isinstance(conteudo, Image.Image):
                conteudo.save(caminho_dest, format="PNG")
            else:
                data = conteudo.getvalue() if hasattr(conteudo, "getvalue") else conteudo
                with open(caminho_dest, "wb") as f: f.write(data)
            
            evid_meta[marcador].append({
                "name": item["name"],
                "file": f"evidencias/{fname}",
                "type": item["type"]
            })
            
    estado = {
        "form_state": {k: st.session_state.get(k) for k in FORM_KEYS},
        "evidencias": evid_meta
    }
    
    with open(pasta / "estado.json", "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)
    
    st.session_state.relatorio_atual = nome_norm
    st.toast(f"Progresso de '{nome}' salvo!")

def carregar_relatorio(nome_pasta):
    pasta = BASE_RELATORIOS_DIR / nome_pasta
    with open(pasta / "estado.json", "r", encoding="utf-8") as f:
        estado = json.load(f)
    
    for k, v in estado.get("form_state", {}).items():
        st.session_state[k] = v
        
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
    for marcador, lista in estado.get("evidencias", {}).items():
        for meta in lista:
            p = pasta / meta["file"]
            if p.exists():
                with open(p, "rb") as f:
                    bio = io.BytesIO(f.read())
                    bio.name = meta["name"]
                    st.session_state.dados_sessao[marcador].append({
                        "name": meta["name"],
                        "content": bio,
                        "type": meta["type"]
                    })
    st.session_state.relatorio_atual = nome_pasta
    st.rerun()

def excluir_relatorio(nome_pasta):
    shutil.rmtree(BASE_RELATORIOS_DIR / nome_pasta)
    st.session_state.relatorio_atual = ""
    st.rerun()

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3208/3208726.png", width=100)
    st.title("Painel de Controle")
    st.markdown("---")
    
    total_anexos = sum(len(v) for v in st.session_state.dados_sessao.values())
    st.metric("Total de Anexos", total_anexos)
    
    if st.button("🗑 Limpar Todos os Dados", use_container_width=True):
        st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
        st.rerun()

# --- FUNÇÕES CORE ---
def excel_para_imagem(doc_template, arquivo_excel):
    try:
        if hasattr(arquivo_excel, 'seek'): arquivo_excel.seek(0)
        df = pd.read_excel(arquivo_excel, sheet_name="TRANSFERENCIAS", usecols=[3, 4], skiprows=2, nrows=14, header=None)
        df = df.fillna('')
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.axis('off')
        tabela = ax.table(cellText=df.values, loc='center', cellLoc='center', colWidths=[0.45, 0.45])
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(11)
        tabela.scale(1.2, 1.8)
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=200)
        plt.close(fig)
        img_buf.seek(0)
        return InlineImage(doc_template, img_buf, width=Mm(DIMENSOES_CAMPOS["TABELA_TRANSFERENCIA"]))
    except Exception as e:
        st.error(f"Erro Excel: {e}")
        return None

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
        if isinstance(item, Image.Image):
            img_buf = io.BytesIO()
            item.save(img_buf, format='PNG')
            img_buf.seek(0)
            return [InlineImage(doc_template, img_buf, width=Mm(largura))]
        
        if isinstance(item, bytes):
            return [InlineImage(doc_template, io.BytesIO(item), width=Mm(largura))]
            
        if hasattr(item, 'seek'): item.seek(0)
        
        ext = getattr(item, 'name', '').lower()
        if marcador == "TABELA_TRANSFERENCIA" and (ext.endswith(".xlsx") or ext.endswith(".xls")):
            res = excel_para_imagem(doc_template, item)
            return [res] if res else []
            
        if ext.endswith(".pdf"):
            pdf = fitz.open(stream=item.read(), filetype="pdf")
            imgs = []
            for pg in pdf:
                pix = pg.get_pixmap(matrix=fitz.Matrix(2, 2))
                imgs.append(InlineImage(doc_template, io.BytesIO(pix.tobytes()), width=Mm(largura)))
            pdf.close()
            return imgs
            
        return [InlineImage(doc_template, item, width=Mm(largura))]
    except Exception:
        return []

# --- UI PRINCIPAL ---
st.title("Automação de Relatórios - UPA Pacheco")
st.caption("Versão 0.7.12")

# --- NOVO: GESTOR DE RELATÓRIOS (v0.7.12) ---
with st.expander("📂 Gestor de Relatórios Guardados", expanded=not st.session_state.relatorio_atual):
    col_g1, col_g2 = st.columns([2, 1])
    with col_g1:
        lista_pastas = [p.name for p in BASE_RELATORIOS_DIR.iterdir() if p.is_dir()]
        sel_disco = st.selectbox("Relatórios Guardados", ["-- Selecionar --"] + lista_pastas)
        ca1, ca2 = st.columns(2)
        if ca1.button("📥 Carregar Selecionado", use_container_width=True) and sel_disco != "-- Selecionar --":
            carregar_relatorio(sel_disco)
        if ca2.button("🗑️ Excluir Selecionado", use_container_width=True) and sel_disco != "-- Selecionar --":
            excluir_relatorio(sel_disco)
    with col_g2:
        novo_nome = st.text_input("Nome do Relatório", placeholder="Ex: Pacheco_Marco_2025")
        if st.button("💾 Salvar Progresso", use_container_width=True, type="primary"):
            salvar_relatorio(novo_nome)

st.markdown("---")

t_manual, t_evidencia = st.tabs(["Dados", "Evidências"])

with t_manual:
    st.markdown("### Configuração do Período e Metas")
    
    c1, c2, c3 = st.columns(3)
    meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    with c1: 
        mes_selecionado = st.selectbox("Mês de Referência", meses_pt, key="sel_mes")
    with c2: 
        ano_selecionado = st.selectbox("Ano", [2024, 2025, 2026, 2027], index=2, key="sel_ano")
    with c3:
        st.text_input("Total de Atendimentos", key="in_total")

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
    with c16: st.text_input("Total de Transferências", key="in_tt")
    with c17: st.text_input("Total de Óbitos", key="in_to")
    with c18: st.text_input("Óbito < 24h", key="in_to_menor")

    c19, c20, c21 = st.columns(3)
    with c19: st.text_input("Óbito > 24h", key="in_to_maior")

with t_evidencia:
    labels = {
        "IMAGEM_PRINT_ATENDIMENTO": "Prints Atendimento", "PRINT_CLASSIFICACAO": "Classificação de Risco", 
        "IMAGEM_DOCUMENTO_RAIO_X": "Doc. Raio-X", "TABELA_TRANSFERENCIA": "Tabela Transferência", 
        "GRAFICO_TRANSFERENCIA": "Gráfico Transferência", "TABELA_OBITO": "Tab. Óbito", 
        "TABELA_CCIH": "Tabela CCIH", "TABELA_QUALITATIVA_IMG": "Tab. Qualitativa",
        "IMAGEM_NEP": "Imagens NEP", "IMAGEM_MELHORIAS": "Melhorias",
        "GRAFICO_OUVIDORIA": "Gráfico Ouvidoria", "PDF_OUVIDORIA_INTERNA": "Tabela Ouvidoria"
    }
    
    blocos = [
        ["IMAGEM_PRINT_ATENDIMENTO", "PRINT_CLASSIFICACAO", "IMAGEM_DOCUMENTO_RAIO_X"],
        ["TABELA_TRANSFERENCIA", "GRAFICO_TRANSFERENCIA"],
        ["TABELA_OBITO", "TABELA_CCIH", "TABELA_QUALITATIVA_IMG"],
        ["IMAGEM_NEP", "IMAGEM_MELHORIAS", "GRAFICO_OUVIDORIA", "PDF_OUVIDORIA_INTERNA"]
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
                    if pasted is not None and pasted.image_data is not None:
                        st.session_state.dados_sessao[m].append({"name": f"Captura_{len(st.session_state.dados_sessao[m]) + 1}.png", "content": pasted.image_data, "type": "p"})
                        st.toast(f"Anexado em: {labels[m]}")
                        time.sleep(0.5)
                        st.rerun()
                with cb:
                    f_up = st.file_uploader("Upload", type=['png', 'jpg', 'pdf', 'xlsx'], key=f"f_{m}_{b_idx}", label_visibility="collapsed")
                    if f_up:
                        if f_up.name not in [x['name'] for x in st.session_state.dados_sessao[m]]:
                            st.session_state.dados_sessao[m].append({"name": f_up.name, "content": f_up, "type": "f"})
                            st.rerun()

                if st.session_state.dados_sessao[m]:
                    for i_idx, item in enumerate(st.session_state.dados_sessao[m]):
                        with st.expander(f"{item['name']}", expanded=False):
                            is_image = item['type'] == "p" or item['name'].lower().endswith(('.png', '.jpg', '.jpeg'))
                            if is_image:
                                st.image(item['content'], width='stretch')
                            else:
                                st.info(f"Ficheiro {item['name'].split('.')[-1].upper()} pronto para o relatório.")
                                
                            if st.button("Remover", key=f"del_{m}_{i_idx}_{b_idx}"):
                                st.session_state.dados_sessao[m].pop(i_idx)
                                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# --- GERAÇÃO FINAL COM PROTEÇÃO (v0.7.12) ---
if st.button("🚀 FINALIZAR E GERAR RELATÓRIO", type="primary", use_container_width=True):
    try:
        progress_bar = st.progress(0)
        with tempfile.TemporaryDirectory() as tmp:
            docx_p = os.path.join(tmp, "relatorio.docx")
            doc = DocxTemplate("template-upa-pacheco.docx")
            
            mes_ano_ref = f"{st.session_state.get('sel_mes', 'N/D')}/{st.session_state.get('sel_ano', 'N/D')}"
            
            # Geração Flexível: .get(chave, valor_padrao)
            dados_finais = {
                "SISTEMA_MES_REFERENCIA": mes_ano_ref,
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

            for marcador in DIMENSOES_CAMPOS.keys():
                lista_imgs = []
                for item in st.session_state.dados_sessao[marcador]:
                    res = processar_item_lista(doc, item['content'], marcador)
                    if res: lista_imgs.extend(res)
                dados_finais[marcador] = lista_imgs
            
            doc.render(dados_finais)
            doc.save(docx_p)
            progress_bar.progress(60)
            
            st.success("✅ Relatório gerado!")
            c_down1, c_down2 = st.columns(2)
            with c_down1:
                with open(docx_p, "rb") as f_w:
                    st.download_button(label="Baixar WORD (.docx)", data=f_w.read(), file_name=f"RELATÓRIO ASSISTENCIAL MENSAL - UPA PACHECO_{mes_ano_ref}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            with c_down2:
                try:
                    converter_para_pdf(docx_p, tmp)
                    pdf_p = os.path.join(tmp, "relatorio.pdf")
                    if os.path.exists(pdf_p):
                        with open(pdf_p, "rb") as f_p:
                            st.download_button(label="Baixar PDF", data=f_p.read(), file_name=f"RELATÓRIO ASSISTENCIAL MENSAL - UPA PACHECO_{mes_ano_ref}.pdf", mime="application/pdf", use_container_width=True)
                except: st.warning("LibreOffice não encontrado.")
    except Exception as e: st.error(f"Erro na geração: {e}")

st.caption("Desenvolvido por Leonardo Barcelos Martins")


