import streamlit as st
from docx import Document
from datetime import datetime
import streamlit.components.v1 as components
import html

# Função para extrair etapas e decisões do .docx
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

# Função para gerar BPMN XML com bloco DI (Diagram Interchange)
def gerar_bpmn_xml(etapas):
    xml_header = '''<?xml version="1.0" encoding="UTF-8"?>
<definitions xmlns="http://www.omg.org/spec/BPMN/20100524/MODEL"
             xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
             xmlns:bpmndi="http://www.omg.org/spec/BPMN/20100524/DI"
             xmlns:omgdc="http://www.omg.org/spec/DD/20100524/DC"
             xmlns:omgdi="http://www.omg.org/spec/DD/20100524/DI"
             targetNamespace="http://bpmn.io/schema/bpmn">
  <process id="ProcessoImportadoDocx" isExecutable="true">
    <startEvent id="StartEvent_1" name="Início"/>
'''

    xml_elements = []
    xml_flows = []
    xml_diagram = []
    task_id = 1
    gateway_id = 1
    next_from = "StartEvent_1"
    pos_y = 100
    node_pos = {"StartEvent_1": (100, pos_y)}

    for item in etapas:
        if item["tipo"] == "etapa":
            tid = f"Task_{task_id}"
            xml_elements.append(f'    <task id="{tid}" name="{item["nome"]}" />')
            xml_flows.append((next_from, tid))
            node_pos[tid] = (100 + 150 * task_id, pos_y)
            next_from = tid
            task_id += 1

        elif item["tipo"] == "gateway":
            gid = f"Gateway_{gateway_id}"
            tid_sim = f"Task_{task_id}"
            tid_nao = f"Task_{task_id+1}"
            end_sim = f"EndEvent_{gateway_id}_1"
            end_nao = f"EndEvent_{gateway_id}_2"

            xml_elements.append(f'    <exclusiveGateway id="{gid}" name="{item["condicao"]}" />')
            xml_elements.append(f'    <task id="{tid_sim}" name="{item["sim"]}" />')
            xml_elements.append(f'    <task id="{tid_nao}" name="{item["nao"]}" />')
            xml_elements.append(f'    <endEvent id="{end_sim}" name="Fim"/>')
            xml_elements.append(f'    <endEvent id="{end_nao}" name="Fim"/>')

            xml_flows.append((next_from, gid))
            xml_flows.append((gid, tid_sim, "sim"))
            xml_flows.append((gid, tid_nao, "não"))
            xml_flows.append((tid_sim, end_sim))
            xml_flows.append((tid_nao, end_nao))

            node_pos[gid] = (100 + 150 * task_id, pos_y)
            node_pos[tid_sim] = (100 + 150 * (task_id + 1), pos_y - 50)
            node_pos[tid_nao] = (100 + 150 * (task_id + 1), pos_y + 50)
            node_pos[end_sim] = (100 + 150 * (task_id + 2), pos_y - 50)
            node_pos[end_nao] = (100 + 150 * (task_id + 2), pos_y + 50)

            next_from = None
            task_id += 2
            gateway_id += 1

    if next_from not in [None, "None"]:
        eid = "EndEvent_final"
        xml_elements.append(f'    <endEvent id="{eid}" name="Fim"/>')
        xml_flows.append((next_from, eid))
        node_pos[eid] = (100 + 150 * task_id, pos_y)

    xml_body = "\n".join(xml_elements)

    flow_xml = ""
    for i, item in enumerate(xml_flows):
        fid = f"Flow_{i+1}"
        if len(item) == 2:
            s, t = item
            flow_xml += f'    <sequenceFlow id="{fid}" sourceRef="{s}" targetRef="{t}"/>\n'
        else:
            s, t, cond = item
            flow_xml += f'''    <sequenceFlow id="{fid}" sourceRef="{s}" targetRef="{t}">
      <conditionExpression xsi:type="tFormalExpression"><![CDATA[{cond}]]></conditionExpression>
    </sequenceFlow>\n'''

    xml_end = "  </process>\n"

    xml_end += "  <bpmndi:BPMNDiagram id=\"BPMNDiagram_1\">\n    <bpmndi:BPMNPlane id=\"BPMNPlane_1\" bpmnElement=\"ProcessoImportadoDocx\">\n"
    for elem_id, (x, y) in node_pos.items():
        shape = "task"
        if "StartEvent" in elem_id:
            shape = "startEvent"
        elif "EndEvent" in elem_id:
            shape = "endEvent"
        elif "Gateway" in elem_id:
            shape = "exclusiveGateway"

        xml_end += f'''      <bpmndi:BPMNShape id="{elem_id}_di" bpmnElement="{elem_id}">
        <omgdc:Bounds x="{x}" y="{y}" width="100" height="80" />
      </bpmndi:BPMNShape>\n'''

    xml_end += "    </bpmndi:BPMNPlane>\n  </bpmndi:BPMNDiagram>\n</definitions>"

    return xml_header + xml_body + "\n" + flow_xml + xml_end

# Streamlit App
st.set_page_config(page_title="POP para BPMN", layout="centered")
st.title("📄 Conversor de Procedimento Operacional para BPMN")

uploaded_file = st.file_uploader("Envie um arquivo .docx com o procedimento:", type="docx")

if uploaded_file:
    etapas = extrair_etapas_e_decisoes(uploaded_file)

    st.subheader("🔍 Etapas e Decisões Extraídas")
    for etapa in etapas:
        st.json(etapa)

    if st.button("🔁 Gerar Arquivo BPMN"):
        xml_bpmn = gerar_bpmn_xml(etapas)
        filename = f"fluxograma_{datetime.now().strftime('%Y%m%d%H%M%S')}.bpmn"
        st.download_button("📅 Baixar BPMN", xml_bpmn, file_name=filename, mime="application/xml")

        st.subheader("📊 Visualização do Fluxograma")
        bpmn_escaped = html.escape(xml_bpmn).replace("\n", "").replace("'", "\\'")
        bpmn_html = f"""
        <!DOCTYPE html>
        <html>
          <head>
            <script src='https://unpkg.com/bpmn-js@11.5.0/dist/bpmn-viewer.development.js'></script>
            <style>
              html, body {{ margin: 0; padding: 0; height: 100%; }}
              #canvas {{ height: 600px; border: 1px solid #ccc; background-color: #f8f9fa; }}
            </style>
          </head>
          <body>
            <div id='canvas'></div>
            <script>
              const bpmnXML = `{bpmn_escaped}`;
              const viewer = new BpmnJS({{ container: '#canvas' }});
              viewer.importXML(bpmnXML).then(() => {{
                viewer.get('canvas').zoom('fit-viewport');
              }}).catch(err => {{
                document.body.innerText = 'Erro ao carregar BPMN: ' + err;
              }});
            </script>
          </body>
        </html>
        """
        components.html(bpmn_html, height=650, scrolling=True)

        st.markdown("""
        🔗 Caso prefira, abra manualmente o arquivo BPMN em:
        [https://bpmn.io/toolkit/bpmn-js/demo/modeler.html](https://bpmn.io/toolkit/bpmn-js/demo/modeler.html)
        """)
