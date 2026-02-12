import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
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
import base64

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios V0.7.10", layout="wide")

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
    .stRadio > div { flex-direction: row; gap: 20px; }
    </style>
    """, unsafe_allow_html=True)

# --- DICIONÁRIO DE DIMENSÕES ---
DIMENSOES_CAMPOS = {
    "IMAGEM_PRINT_ATENDIMENTO": 165, "PRINT_CLASSIFICACAO": 160,
    "IMAGEM_DOCUMENTO_RAIO_X": 165, "TABELA_TRANSFERENCIA": 90,
    "GRAFICO_TRANSFERENCIA": 160, "TABELA_OBITO": 180, 
    "TABELA_CCIH": 180, "IMAGEM_NEP": 160,
    "IMAGEM_TREINAMENTO_INTERNO": 160, "IMAGEM_MELHORIAS": 160,
    "GRAFICO_OUVIDORIA": 155, "PDF_OUVIDORIA_INTERNA": 165,
    "TABELA_QUALITATIVA_IMG": 160
}

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}

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
        if hasattr(item, 'seek'): item.seek(0)
        if isinstance(item, bytes):
            return [InlineImage(doc_template, io.BytesIO(item), width=Mm(largura))]
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
    except Exception: return []

# --- FUNÇÕES DE IMPORTAÇÃO/EXPORTAÇÃO ---
def exportar_pacote():
    pacote = {}
    for marcador, itens in st.session_state.dados_sessao.items():
        pacote[marcador] = []
        for it in itens:
            content = it['content']
            if hasattr(content, 'getvalue'):
                b64 = base64.b64encode(content.getvalue()).decode()
            elif isinstance(content, bytes):
                b64 = base64.b64encode(content).decode()
            else:
                continue
            pacote[marcador].append({"name": it['name'], "type": it['type'], "content": b64})
    return json.dumps(pacote)

def importar_pacote(json_str):
    try:
        pacote = json.loads(json_str)
        for marcador, itens in pacote.items():
            st.session_state.dados_sessao[marcador] = []
            for it in itens:
                raw_bytes = base64.b64decode(it['content'])
                st.session_state.dados_sessao[marcador].append({
                    "name": it['name'],
                    "type": it['type'],
                    "content": io.BytesIO(raw_bytes)
                })
        st.toast("✅ Dados da Unidade importados com sucesso!")
        time.sleep(1)
        st.rerun()
    except Exception as e:
        st.error(f"Erro na importação: {e}")

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3208/3208726.png", width=100)
    st.title("Sistema de Gestão")
    st.markdown("---")
    modo = st.radio("Selecione o Perfil:", ["Analista (Relatório)", "Unidade (Envio de Dados)"])
    st.markdown("---")
    if st.button("🗑️ Limpar Sessão", width='stretch'):
        st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}
        st.rerun()

# --- INTERFACE UNIDADE (PONTA) ---
if modo == "Unidade (Envio de Dados)":
    st.title("🚀 Portal de Envio - Unidade")
    st.info("Unidade: Use esta aba para carregar todos os documentos do mês. Ao terminar, clique em 'Gerar Pacote de Dados'.")
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        u_mes = st.selectbox("Mês de Referência", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"], key="u_mes")
    with col_u2:
        u_ano = st.selectbox("Ano", [2024, 2025, 2026, 2027], index=2, key="u_ano")

    st.markdown("### 📁 Carregamento de Evidências")
    labels = {
        "IMAGEM_PRINT_ATENDIMENTO": "Prints Atendimento", "PRINT_CLASSIFICACAO": "Classificação de Risco", 
        "IMAGEM_DOCUMENTO_RAIO_X": "Doc. Raio-X", "TABELA_TRANSFERENCIA": "Planilha de Transferências (Excel)", 
        "GRAFICO_TRANSFERENCIA": "Gráfico Transferência", "TABELA_OBITO": "Tab. Óbito", 
        "TABELA_CCIH": "Tabela CCIH", "TABELA_QUALITATIVA_IMG": "Metas Qualitativas",
        "IMAGEM_NEP": "Fotos NEP", "IMAGEM_TREINAMENTO_INTERNO": "Treinamentos", 
        "IMAGEM_MELHORIAS": "Melhorias Realizadas", "GRAFICO_OUVIDORIA": "Gráfico Ouvidoria", "PDF_OUVIDORIA_INTERNA": "Relatório Ouvidoria"
    }

    # Layout simplificado para a unidade
    for m, label in labels.items():
        with st.expander(f"➕ {label}", expanded=False):
            f_up = st.file_uploader(f"Anexar {label}", type=['png', 'jpg', 'pdf', 'xlsx'], key=f"u_f_{m}")
            if f_up:
                if f_up.name not in [x['name'] for x in st.session_state.dados_sessao[m]]:
                    st.session_state.dados_sessao[m].append({"name": f_up.name, "content": f_up, "type": "f"})
                    st.toast(f"Anexado: {f_up.name}")
            
            if st.session_state.dados_sessao[m]:
                for i_idx, item in enumerate(st.session_state.dados_sessao[m]):
                    st.caption(f"✅ {item['name']}")
                    if st.button(f"Remover {i_idx}", key=f"u_del_{m}_{i_idx}"):
                        st.session_state.dados_sessao[m].pop(i_idx)
                        st.rerun()

    st.markdown("---")
    if st.button("📦 GERAR PACOTE DE DADOS (.tatico)", type="primary", width='stretch'):
        pacote_json = exportar_pacote()
        st.download_button(
            label="⬇️ BAIXAR PACOTE PARA O ANALISTA",
            data=pacote_json,
            file_name=f"DADOS_UNIDADE_{u_mes}_{u_ano}.tatico",
            mime="application/json",
            width='stretch'
        )

# --- INTERFACE ANALISTA (RELATÓRIO) ---
else:
    st.title("📊 Automação de Relatórios - Analista")
    
    # Opção de Importação
    with st.expander("📥 Importar Dados da Unidade", expanded=True):
        f_import = st.file_uploader("Arraste o arquivo .tatico enviado pela unidade aqui", type=['tatico'])
        if f_import:
            if st.button("Confirmar Importação de Dados"):
                importar_pacote(f_import.read().decode())

    t_manual, t_evidencia = st.tabs(["📝 Dados", "📁 Evidências"])

    with t_manual:
        st.markdown("### 📅 Período e Metas")
        c1, c2, c3 = st.columns(3)
        meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        with c1: mes_sel = st.selectbox("Mês de Referência", meses_pt, key="sel_mes")
        with c2: ano_sel = st.selectbox("Ano", [2024, 2025, 2026, 2027], index=2, key="sel_ano")
        with c3: st.text_input("Total de Atendimentos", key="in_total")

        mes_num = meses_pt.index(mes_sel) + 1
        dias_no_mes = calendar.monthrange(ano_sel, mes_num)[1]
        meta_calc = dias_no_mes * META_DIARIA_CONTRATO
        
        c4, c5, c6 = st.columns(3)
        with c4: st.text_input("Meta do Mês (Calculada)", value=str(meta_calc), disabled=True)
        with c5: st.text_input("Meta -25% (Calculada)", value=str(int(meta_calc*0.75)), disabled=True)
        with c6: st.text_input("Meta +25% (Calculada)", value=str(int(meta_calc*1.25)), disabled=True)

        st.markdown("---")
        st.markdown("### 🏥 Dados Assistenciais")
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
        labels_ev = {
            "IMAGEM_PRINT_ATENDIMENTO": "Prints Atendimento", "PRINT_CLASSIFICACAO": "Classificação de Risco", 
            "IMAGEM_DOCUMENTO_RAIO_X": "Doc. Raio-X", "TABELA_TRANSFERENCIA": "Tabela Transferência", 
            "GRAFICO_TRANSFERENCIA": "Gráfico Transferência", "TABELA_OBITO": "Tab. Óbito", 
            "TABELA_CCIH": "Tabela CCIH", "TABELA_QUALITATIVA_IMG": "Tab. Qualitativa",
            "IMAGEM_NEP": "Imagens NEP", "IMAGEM_TREINAMENTO_INTERNO": "Treinamento Interno", 
            "IMAGEM_MELHORIAS": "Melhorias", "GRAFICO_OUVIDORIA": "Gráfico Ouvidoria", "PDF_OUVIDORIA_INTERNA": "Relatório Ouvidoria"
        }
        
        blocos_ev = [
            ["IMAGEM_PRINT_ATENDIMENTO", "PRINT_CLASSIFICACAO", "IMAGEM_DOCUMENTO_RAIO_X"],
            ["TABELA_TRANSFERENCIA", "GRAFICO_TRANSFERENCIA"],
            ["TABELA_OBITO", "TABELA_CCIH", "TABELA_QUALITATIVA_IMG"],
            ["IMAGEM_NEP", "IMAGEM_TREINAMENTO_INTERNO", "IMAGEM_MELHORIAS", "GRAFICO_OUVIDORIA", "PDF_OUVIDORIA_INTERNA"]
        ]

        for b_idx, lista_m in enumerate(blocos_ev):
            st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
            col_esq, col_dir = st.columns(2)
            for idx, m in enumerate(lista_m):
                target = col_esq if idx % 2 == 0 else col_dir
                with target:
                    st.markdown(f"<span class='upload-label'>{labels_ev.get(m, m)}</span>", unsafe_allow_html=True)
                    ca, cb = st.columns([1, 1])
                    with ca:
                        key_p = f"p_{m}_{len(st.session_state.dados_sessao[m])}"
                        pasted = paste_image_button(label="Colar Print", key=key_p)
                        if pasted is not None and pasted.image_data is not None:
                            st.session_state.dados_sessao[m].append({"name": f"Captura_{len(st.session_state.dados_sessao[m]) + 1}.png", "content": pasted.image_data, "type": "p"})
                            st.toast(f"📸 Anexado em: {labels_ev[m]}")
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
                            with st.expander(f"📄 {item['name']}", expanded=False):
                                is_img = item['type'] == "p" or item['name'].lower().endswith(('.png', '.jpg', '.jpeg'))
                                if is_img:
                                    st.image(item['content'], width='stretch')
                                else:
                                    st.info(f"Visualização indisponível ({item['name'].split('.')[-1].upper()}).")
                                if st.button("Remover", key=f"del_{m}_{i_idx}_{b_idx}"):
                                    st.session_state.dados_sessao[m].pop(i_idx)
                                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    # --- GERAÇÃO FINAL ---
    if st.button("🚀 FINALIZAR E GERAR RELATÓRIO", type="primary", width='stretch'):
        try:
            with st.spinner("Gerando documentos..."):
                with tempfile.TemporaryDirectory() as tmp:
                    docx_p = os.path.join(tmp, "relatorio.docx")
                    doc = DocxTemplate("template.docx")
                    
                    dados_finais = {
                        "SISTEMA_MES_REFERENCIA": f"{mes_sel}/{ano_sel}",
                        "ANALISTA_TOTAL_ATENDIMENTOS": st.session_state.get("in_total", ""),
                        "TOTAL_RAIO_X": st.session_state.get("in_rx", ""),
                        "ANALISTA_META_MES": str(meta_calc),
                        "ANALISTA_META_MINUS_25": str(int(meta_calc*0.75)),
                        "ANALISTA_META_PLUS_25": str(int(meta_calc*1.25)),
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
                    
                    st.success("✅ Relatório pronto para download!")
                    cd1, cd2 = st.columns(2)
                    with cd1:
                        with open(docx_p, "rb") as f_w:
                            st.download_button("📥 WORD (.docx)", f_w.read(), f"Relatorio_{mes_sel}.docx", width='stretch')
                    with cd2:
                        try:
                            converter_para_pdf(docx_p, tmp)
                            pdf_p = os.path.join(tmp, "relatorio.pdf")
                            if os.path.exists(pdf_p):
                                with open(pdf_p, "rb") as f_p:
                                    st.download_button("📥 PDF", f_p.read(), f"Relatorio_{mes_sel}.pdf", width='stretch')
                        except: st.warning("LibreOffice não encontrado.")
        except Exception as e: st.error(f"Erro: {e}")

st.caption("Desenvolvido por Leonardo Barcelos Martins")
