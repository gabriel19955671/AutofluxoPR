import streamlit as st
from docx import Document
from datetime import datetime
import streamlit.components.v1 as components
import html
import xml.etree.ElementTree as ET

# Função para extrair etapas e gateways do documento
def extrair_etapas_e_decisoes(docx_file):
    doc = Document(docx_file)
    etapas = []
    condicional = None
    for para in doc.paragraphs:
        texto = para.text.strip()
        if not texto:
            continue
        if texto.lower().startswith("etapa"):
            etapas.append({"tipo": "etapa", "nome": texto.replace("Etapa:", "").strip()})
        elif texto.lower().startswith("se"):
            condicional = {"tipo": "gateway", "condicao": texto.replace("Se", "").replace(":", "").strip(), "sim": None, "nao": None}
        elif texto.lower().startswith("senão") or texto.lower().startswith("senao"):
            continue
        elif condicional and condicional["sim"] is None:
            condicional["sim"] = texto.replace("Etapa:", "").strip()
        elif condicional and condicional["nao"] is None:
            condicional["nao"] = texto.replace("Etapa:", "").strip()
            etapas.append(condicional)
            condicional = None
    return etapas

# Função para gerar XML no formato draw.io

def gerar_drawio_xml(etapas):
    root = ET.Element("mxGraphModel")
    root.set("dx", "1000")
    root.set("dy", "1000")
    root.set("grid", "1")
    root.set("gridSize", "10")
    root.set("guides", "1")
    root.set("tooltips", "1")
    root.set("connect", "1")
    root.set("arrows", "1")
    root.set("fold", "1")
    root.set("page", "1")
    root.set("pageScale", "1")
    root.set("pageWidth", "850")
    root.set("pageHeight", "1100")
    root.set("math", "0")
    root.set("shadow", "0")

    root_cell = ET.SubElement(ET.SubElement(root, "root"), "mxCell", id="0")
    ET.SubElement(root.find("root"), "mxCell", id="1", parent="0")

    y = 40
    step_id = 2
    last_id = "1"

    for etapa in etapas:
        nome = etapa.get("nome") or etapa.get("condicao") or "Etapa"
        style = "shape=process;whiteSpace=wrap;html=1;" if etapa["tipo"] == "etapa" else "shape=rhombus;whiteSpace=wrap;html=1;"
        step = ET.SubElement(root.find("root"), "mxCell", id=str(step_id), value=nome, style=style, vertex="1", parent="1")
        ET.SubElement(step, "mxGeometry", x="100", y=str(y), width="160", height="60", as_="geometry")

        # Conector
        if step_id > 2:
            edge = ET.SubElement(root.find("root"), "mxCell", id=str(step_id+1000), style="endArrow=block;", edge="1", parent="1", source=str(step_id-1), target=str(step_id))
            ET.SubElement(edge, "mxGeometry", relative="1", as_="geometry")

        y += 100
        last_id = str(step_id)
        step_id += 1

    xml_str = ET.tostring(root, encoding="utf-8", method="xml").decode("utf-8")
    return f'<?xml version="1.0" encoding="UTF-8"?>\n{xml_str}'

# Interface Streamlit
st.set_page_config(page_title="POP para Fluxo", layout="centered")
st.title("📄 Conversor de Procedimento Operacional para Fluxograma")

uploaded_file = st.file_uploader("Envie um arquivo .docx com o procedimento:", type="docx")

if uploaded_file:
    etapas = extrair_etapas_e_decisoes(uploaded_file)
    st.subheader("🔍 Etapas e Decisões Extraídas")
    for etapa in etapas:
        st.json(etapa)

    tipo_fluxo = st.radio("Escolha o formato do fluxograma:", ["BPMN", "Draw.io"])

    if st.button("🔁 Gerar Arquivo de Fluxo"):
        if tipo_fluxo == "BPMN":
            st.warning("🚧 A geração de BPMN está temporariamente desativada nesta versão.")
        else:
            drawio_xml = gerar_drawio_xml(etapas)
            filename = f"fluxograma_{datetime.now().strftime('%Y%m%d%H%M%S')}.xml"
            st.download_button("📥 Baixar Fluxograma (Draw.io)", drawio_xml, file_name=filename, mime="application/xml")
            st.success("✅ Arquivo gerado com sucesso. Importe em https://app.diagrams.net")
