import streamlit as st
from docx import Document
from datetime import datetime
import streamlit.components.v1 as components
import html  # Novo para escapar XML

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

# Função para gerar BPMN XML
def gerar_bpmn_xml(etapas):
    xml = '''<?xml version="1.0" encoding="UTF-8"?>
<definitions xmlns="http://www.omg.org/spec/BPMN/20100524/MODEL"
             xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
             id="Definitions_1"
             targetNamespace="http://bpmn.io/schema/bpmn">
  <process id="ProcessoImportadoDocx" isExecutable="true">
    <startEvent id="StartEvent_1" name="Início"/>
'''
    seq = []
    task_id = 1
    gateway_id = 1
    next_from = "StartEvent_1"

    for item in etapas:
        if item["tipo"] == "etapa":
            tid = f"Task_{task_id}"
            xml += f'    <task id="{tid}" name="{item["nome"]}" />\n'
            seq.append((next_from, tid))
            next_from = tid
            task_id += 1

        elif item["tipo"] == "gateway":
            gid = f"Gateway_{gateway_id}"
            tid_sim = f"Task_{task_id}"
            tid_nao = f"Task_{task_id+1}"

            xml += f'    <exclusiveGateway id="{gid}" name="{item["condicao"]}" />\n'
            xml += f'    <task id="{tid_sim}" name="{item["sim"]}" />\n'
            xml += f'    <task id="{tid_nao}" name="{item["nao"]}" />\n'

            seq.append((next_from, gid))
            seq.append((gid, tid_sim, "sim"))
            seq.append((gid, tid_nao, "não"))

            seq.append((tid_sim, f"EndEvent_{gateway_id}_1"))
            seq.append((tid_nao, f"EndEvent_{gateway_id}_2"))

            xml += f'    <endEvent id="EndEvent_{gateway_id}_1" name="Fim"/>\n'
            xml += f'    <endEvent id="EndEvent_{gateway_id}_2" name="Fim"/>\n'

            next_from = None  # fim de ramificação
            task_id += 2
            gateway_id += 1

    if next_from:
        xml += f'    <endEvent id="EndEvent_final" name="Fim"/>\n'
        seq.append((next_from, "EndEvent_final"))

    flow_count = 1
    for item in seq:
        if len(item) == 2:
            source, target = item
            xml += f'    <sequenceFlow id="Flow_{flow_count}" sourceRef="{source}" targetRef="{target}"/>\n'
        else:
            source, target, cond = item
            xml += f'''    <sequenceFlow id="Flow_{flow_count}" sourceRef="{source}" targetRef="{target}">
      <conditionExpression xsi:type="tFormalExpression"><![CDATA[{cond}]]></conditionExpression>
    </sequenceFlow>\n'''
        flow_count += 1

    xml += '  </process>\n</definitions>'
    return xml

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
