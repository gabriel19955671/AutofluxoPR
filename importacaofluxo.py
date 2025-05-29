import streamlit as st
from docx import Document
from datetime import datetime
import xml.etree.ElementTree as ET
import pandas as pd
import base64
import tempfile
import os

# Função para extrair estrutura do documento
def extrair_pop_struct(docx_file):
    doc = Document(docx_file)
    etapas = []
    etapa_atual = {}
    for para in doc.paragraphs:
        texto = para.text.strip()
        if not texto:
            continue
        if texto.startswith("[ETAPA]"):
            if etapa_atual:
                etapas.append(etapa_atual)
            etapa_atual = {"Etapa": texto.replace("[ETAPA]", "").strip(), "Responsável": "", "Condição": "", "Sim": "", "Não": ""}
        elif texto.startswith("[RESPONSÁVEL]"):
            etapa_atual["Responsável"] = texto.replace("[RESPONSÁVEL]", "").strip()
        elif texto.startswith("[SE]"):
            etapa_atual["Condição"] = texto.replace("[SE]", "").strip()
        elif texto.startswith("[SIM]"):
            etapa_atual["Sim"] = texto.replace("[SIM]", "").strip()
        elif texto.startswith("[NÃO]") or texto.startswith("[NÃO]"):
            etapa_atual["Não"] = texto.replace("[NÃO]", "").strip()
    if etapa_atual:
        etapas.append(etapa_atual)
    return etapas

# Gera um XML draw.io com conexões lógicas baseadas nas condições e separação por responsável
def gerar_drawio_com_lanes(df):
    df.columns = [col.strip() for col in df.columns]
    if "Responsável" not in df.columns:
        raise KeyError("Coluna 'Responsável' não encontrada. Verifique as tags no documento.")

    root = ET.Element("mxGraphModel", attrib={"dx": "1000", "dy": "1000", "grid": "1", "gridSize": "10", "page": "1", "pageScale": "1", "pageWidth": "850", "pageHeight": "1100"})
    root_elem = ET.SubElement(root, "root")
    ET.SubElement(root_elem, "mxCell", id="0")
    ET.SubElement(root_elem, "mxCell", id="1", parent="0")

    lanes = {}
    y_positions = {}
    blocos = {}
    step_id = 2
    lane_x = 100

    for resp in df["Responsável"].unique():
        if not pd.isna(resp):
            lanes[resp] = lane_x
            y_positions[resp] = 40
            lane_x += 300

    for _, row in df.iterrows():
        nome = row["Etapa"] or "Etapa"
        responsavel = row["Responsável"] or "Indefinido"
        x = lanes.get(responsavel, 100)
        y = y_positions.get(responsavel, 40)
        style = "shape=rhombus;whiteSpace=wrap;html=1;" if row["Condição"] else "shape=process;whiteSpace=wrap;html=1;"
        step = ET.SubElement(root_elem, "mxCell", id=str(step_id), value=nome, style=style, vertex="1", parent="1")
        geometry = ET.SubElement(step, "mxGeometry", x=str(x), y=str(y), width="160", height="60")
        geometry.set("as", "geometry")
        blocos[nome] = str(step_id)
        y_positions[responsavel] += 100
        step_id += 1

    for _, row in df.iterrows():
        origem_id = blocos.get(row["Etapa"])
        if row["Sim"] and row["Sim"] in blocos:
            target_id = blocos[row["Sim"]]
            edge = ET.SubElement(root_elem, "mxCell", id=str(step_id+1000), style="endArrow=block;", edge="1", parent="1", source=origem_id, target=target_id)
            edge_geom = ET.SubElement(edge, "mxGeometry", relative="1")
            edge_geom.set("as", "geometry")
        if row["Não"] and row["Não"] in blocos:
            target_id = blocos[row["Não"]]
            edge = ET.SubElement(root_elem, "mxCell", id=str(step_id+2000), style="endArrow=block;dashed=1;", edge="1", parent="1", source=origem_id, target=target_id)
            edge_geom = ET.SubElement(edge, "mxGeometry", relative="1")
            edge_geom.set("as", "geometry")

    xml_str = ET.tostring(root, encoding="utf-8", method="xml").decode("utf-8")
    return f'<?xml version="1.0" encoding="UTF-8"?>\n{xml_str}'

# Gera link de visualização direta como imagem base64 no Streamlit
def gerar_link_imagem(xml_str):
    b64 = base64.b64encode(xml_str.encode()).decode()
    url = f"https://viewer.diagrams.net/?highlight=0000ff&edit=_blank&layers=1&nav=1#R{b64}"
    return url

# Salva como imagem PNG a partir do XML usando headless Chromium (opcional para ambientes compatíveis)
def salvar_xml_como_png(xml_str):
    temp_dir = tempfile.mkdtemp()
    xml_path = os.path.join(temp_dir, "fluxo.xml")
    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(xml_str)
    return xml_path

# Streamlit App
st.set_page_config(page_title="Editor POP para Fluxograma", layout="centered")
st.title("🧭 POP para Fluxo Interativo")

uploaded_file = st.file_uploader("📄 Envie o arquivo de Procedimento Operacional (.docx):", type="docx")

if uploaded_file:
    dados = extrair_pop_struct(uploaded_file)
    if not dados:
        st.warning("⚠️ Nenhuma etapa foi extraída do documento. Verifique se o POP está formatado corretamente com as tags [ETAPA], [RESPONSÁVEL], etc.")
        st.stop()

    df = pd.DataFrame(dados)
    st.subheader("📝 Etapas extraídas do POP")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    if st.button("📥 Gerar Fluxograma (Draw.io)"):
        try:
            xml = gerar_drawio_com_lanes(edited_df)
            filename = f"fluxograma_{datetime.now().strftime('%Y%m%d%H%M%S')}.xml"
            st.download_button("⬇ Baixar XML Draw.io", xml, file_name=filename, mime="application/xml")

            drawio_link = gerar_link_imagem(xml)
            st.markdown(f"[🔍 Visualizar diretamente no Draw.io]({drawio_link})")
            st.success("✅ Arquivo gerado e pronto para visualizar!")
        except KeyError as e:
            st.error(f"❌ Erro: {str(e)}")
