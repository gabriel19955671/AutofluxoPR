import streamlit as st
from docx import Document
from datetime import datetime
import xml.etree.ElementTree as ET
import pandas as pd
import base64
import tempfile
import os

# Gera um XML draw.io com notacao BPMN visual e lanes visuais agrupadas por responsável
def gerar_drawio_com_lanes(df):
    df.columns = [col.strip() for col in df.columns]
    if "Responsável" not in df.columns:
        raise KeyError("Coluna 'Responsável' não encontrada. Verifique as tags no documento.")

    if df["Responsável"].isnull().any() or (df["Responsável"] == "").any():
        raise KeyError("Todas as etapas devem ter um responsável definido antes de gerar o fluxo.")

    root = ET.Element("mxGraphModel", attrib={"dx": "1000", "dy": "1000", "grid": "1", "gridSize": "10", "page": "1", "pageScale": "1", "pageWidth": "1400", "pageHeight": "1100"})
    root_elem = ET.SubElement(root, "root")
    ET.SubElement(root_elem, "mxCell", id="0")
    ET.SubElement(root_elem, "mxCell", id="1", parent="0")

    lanes = {}
    y_positions = {}
    blocos = {}
    step_id = 2
    lane_x = 100
    lane_width = 260

    for idx, resp in enumerate(df["Responsável"].unique()):
        lane_id = f"lane_{idx}"
        lanes[resp] = lane_x
        y_positions[resp] = 40
        lane = ET.SubElement(root_elem, "mxCell", id=lane_id, value=resp, style="swimlane;startSize=20;", vertex="1", parent="1")
        geom = ET.SubElement(lane, "mxGeometry", x=str(lane_x), y="20", width=str(lane_width), height="900")
        geom.set("as", "geometry")
        lane_x += lane_width + 40

    for idx, row in df.iterrows():
        nome = row["Etapa"] or f"Etapa {idx+1}"
        responsavel = row["Responsável"]
        x = lanes.get(responsavel, 100) + 30
        y = y_positions.get(responsavel, 40)

        if idx == 0:
            estilo = "shape=ellipse;perimeter=ellipsePerimeter;whiteSpace=wrap;html=1;"
        elif row["Condição"]:
            estilo = "shape=rhombus;whiteSpace=wrap;html=1;"
            nome = row["Condição"]
        elif idx == len(df)-1:
            estilo = "shape=ellipse;perimeter=ellipsePerimeter;whiteSpace=wrap;html=1;"
        else:
            estilo = "shape=rectangle;rounded=1;whiteSpace=wrap;html=1;"

        step = ET.SubElement(root_elem, "mxCell", id=str(step_id), value=nome, style=estilo, vertex="1", parent="1")
        geometry = ET.SubElement(step, "mxGeometry", x=str(x), y=str(y), width="160", height="60")
        geometry.set("as", "geometry")
        blocos[row["Etapa"]] = str(step_id)
        y_positions[responsavel] += 100
        step_id += 1

    for _, row in df.iterrows():
        origem_id = blocos.get(row["Etapa"])
        if row["Sim"] and row["Sim"] in blocos:
            target_id = blocos[row["Sim"]]
            edge = ET.SubElement(root_elem, "mxCell", id=str(step_id+1000), style="endArrow=block;", edge="1", parent="1", source=origem_id, target=target_id, value="Sim")
            edge_geom = ET.SubElement(edge, "mxGeometry", relative="1")
            edge_geom.set("as", "geometry")
        if row["Não"] and row["Não"] in blocos:
            target_id = blocos[row["Não"]]
            edge = ET.SubElement(root_elem, "mxCell", id=str(step_id+2000), style="endArrow=block;dashed=1;", edge="1", parent="1", source=origem_id, target=target_id, value="Não")
            edge_geom = ET.SubElement(edge, "mxGeometry", relative="1")
            edge_geom.set("as", "geometry")

    xml_str = ET.tostring(root, encoding="utf-8", method="xml").decode("utf-8")
    return f'<?xml version="1.0" encoding="UTF-8"?>\n{xml_str}'

# Gera link de visualização direta como imagem base64 no Streamlit
def gerar_link_imagem(xml_str):
    b64 = base64.b64encode(xml_str.encode()).decode()
    url = f"https://viewer.diagrams.net/?highlight=0000ff&edit=_blank&layers=1&nav=1#R{b64}"
    return url

st.set_page_config(page_title="Editor POP para Fluxograma", layout="centered")
st.title("🧭 POP para Fluxo Interativo")

arquivo = st.file_uploader("📤 Faça upload do POP (.docx)", type=["docx"])
conteudo = ""

if arquivo:
    doc = Document(arquivo)
    linhas = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    df = pd.DataFrame({"Etapa": linhas, "Responsável": "", "Condição": "", "Sim": "", "Não": ""})

    if st.button("✨ Preencher responsáveis automaticamente"):
        df["Responsável"] = "Fiscal"
        st.info("Responsáveis preenchidos como 'Fiscal' para todas as etapas.")

    st.success("✅ Documento carregado. Edite as colunas abaixo para gerar o fluxo.")
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
else:
    st.info("📄 Aguarde o upload de um arquivo .docx com o conteúdo do POP.")
