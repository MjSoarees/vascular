import os
import base64
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

st.set_page_config(
    page_title="Telessaúde UEA - Triagem Vascular",
    page_icon="🩺",
    layout="wide"
)

def validar_cpf(cpf):
    cpf = ''.join(filter(str.isdigit, str(cpf)))
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    digito1 = (soma * 10) % 11
    if digito1 == 10:
        digito1 = 0
    if digito1 != int(cpf[9]):
        return False
    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    digito2 = (soma * 10) % 11
    if digito2 == 10:
        digito2 = 0
    if digito2 != int(cpf[10]):
        return False
    return True

def calcular_idade(data_nasc):
    try:
        nasc = datetime.strptime(data_nasc.strip(), "%d/%m/%Y")
        hoje = datetime.today()
        idade = hoje.year - nasc.year - ((hoje.month, hoje.day) < (nasc.month, nasc.day))
        return str(idade)
    except Exception:
        return ""

def img_to_base64(path):
    if os.path.exists(path):
        with open(path, "rb") as f:
            data = f.read()
        return base64.b64encode(data).decode()
    return ""

def adicionar_cabecalho_e_marca_dagua(doc, logo_marca_dagua, logos_topo):
    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0]
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 1. Adiciona as 3 logos centralizadas no topo do cabeçalho
    for logofile in logos_topo:
        if os.path.exists(logofile):
            try:
                run = header_para.add_run()
                run.add_picture(logofile, width=Inches(0.9))
                header_para.add_run("   ") 
            except Exception:
                pass

    # 2. Adiciona a marca d'água grande de fundo centralizada
    if logo_marca_dagua and os.path.exists(logo_marca_dagua):
        try:
            watermark_p = header.add_paragraph()
            watermark_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_wm = watermark_p.add_run()
            pic = run_wm.add_picture(logo_marca_dagua, width=Inches(5.5))
            
            inline = pic._inline
            
            watermark_xml = parse_xml(f'''
                <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
                           simplePos="0" relativeHeight="251658240" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1">
                    <wp:simplePos x="0" y="0"/>
                    <wp:positionH relativeFrom="page">
                        <wp:align>center</wp:align>
                    </wp:positionH>
                    <wp:positionV relativeFrom="page">
                        <wp:align>center</wp:align>
                    </wp:positionV>
                    <wp:extent cx="5000000" cy="5000000"/>
                    <wp:effectExtent l="0" t="0" r="0" b="0"/>
                    <wp:wrapNone/>
                    <wp:docPr id="999" name="Watermark"/>
                    <wp:cNvGraphicFramePr/>
                    <graphic xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:nvPr><pic:cNvPr id="0" name="Watermark"/><pic:cNvPicPr/></pic:nvPr>
                                <pic:blipFill>
                                    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="{inline.graphic.graphicData.uri}"/>
                                    <a:stretch><a:fillRect/></a:stretch>
                                </pic:blipFill>
                                <pic:spPr>
                                    <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                        <a:off x="0" y="0"/>
                                        <a:ext cx="5000000" cy="5000000"/>
                                    </a:xfrm>
                                    <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"><a:avLst/></a:avLst></a:prstGeom>
                                </pic:spPr>
                            </pic:pic>
                        </graphicData>
                    </graphic>
                </wp:anchor>
            ''')
            inline.getparent().replace(inline, watermark_xml)
        except Exception:
            pass

# Carrega as logos da pasta "fotos" em base64 para a interface web
img1_b64 = img_to_base64("./fotos/Design sem nome (12).png") or img_to_base64("Design sem nome (12).png.png")
# img2_b64 = img_to_base64("./fotos/logo1.png") or img_to_base64("logo1.png")


# --- EXIBIÇÃO DAS 3 LOGOS CENTRALIZADAS E SEMPRE LADO A LADO (PC E MOBILE) ---
logos_html = '<div style="display: flex; justify-content: center; align-items: center; gap: 15px; margin-bottom: 15px;">'
if img1_b64:
    logos_html += f'<img src="data:image/png;base64,{img1_b64}" style="height: 50px; width: auto;" />'
#if img2_b64:
    #logos_html += f'<img src="data:image/png;base64,{img2_b64}" style="height: 50px; width: auto;" />'

st.markdown(logos_html, unsafe_allow_html=True)
# ----------------------------------------------------------------------------

st.markdown("<h3 style='text-align: center; font-size: 1.25rem;'>🩺 Telessaúde UEA · Triagem Vascular</h3>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Preencha os dados abaixo para gerar a solicitação oficial de teleatendimento.</p>", unsafe_allow_html=True)

# Abas organizadas
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "1. Profissional", "2. Paciente", "3. Queixa", "4. Sintomas", 
    "5. Dor", "6. Histórico", "7. Cirurgias e Hábitos", "8. Exames e Finalização"
])

with tab1:
    st.subheader("Profissional Solicitante")
    prof_nome = st.text_input("Nome completo do profissional (não abreviar) *")
    prof_cpf = st.text_input("CPF")
    prof_profissao = st.text_input("Profissão")
    prof_conselho = st.text_input("Registro do conselho")
    prof_rqe = st.text_input("RQE (se médico)")
    prof_municipio = st.text_input("Município")
    prof_email = st.text_input("E-mail")

with tab2:
    st.subheader("Identificação do Paciente")
    pac_nome = st.text_input("Nome completo do paciente *")
    pac_cpf = st.text_input("CPF do paciente")
    pac_nasc = st.text_input("Data de nascimento (DD/MM/AAAA)")
    pac_sexo = st.selectbox("Sexo", ["", "Masculino", "Feminino"])
    pac_ocupacao = st.text_input("Profissão / Ocupação")
    pac_peso = st.text_input("Peso aproximado (kg)")
    pac_altura = st.text_input("Altura (m)")
    pac_pa = st.text_input("P.A. (pressão arterial)")
    
    pac_indigena = st.selectbox("Paciente indígena?", ["Não", "Sim"])
    pac_etnia = st.text_input("Qual etnia?") if pac_indigena == "Sim" else ""

with tab3:
    st.subheader("Queixa Principal")
    q_motivo = st.text_area("Qual é o principal motivo da consulta? *")
    q_cid = st.text_input("CID")
    q_ciap = st.text_input("CIAP")

with tab4:
    st.subheader("Sintomas Vasculares")
    s_empe = st.selectbox("Permanece muito tempo em pé parado?", ["Não", "Sim"])
    s_sentado = st.selectbox("Permanece muito tempo sentado?", ["Não", "Sim"])
    s_dorpernas = st.selectbox("Sente dor nas pernas?", ["Não", "Sim"])
    
    s_incham = st.selectbox("As pernas incham?", ["Não", "Sim"])
    s_inchaco_quando = st.selectbox("Inchaço ocorre mais:", ["Manhã", "Final do dia", "O tempo todo"]) if s_incham == "Sim" else ""
    s_inchaco_onde = st.selectbox("Inchaço é em:", ["Apenas uma perna", "Ambas as pernas"]) if s_incham == "Sim" else ""
    
    s_varizes = st.selectbox("Nota veias dilatadas ou varizes?", ["Não", "Sim"])
    s_pele = st.selectbox("Pele escurecida/manchada?", ["Não", "Sim"])
    
    s_feridas = st.selectbox("Há feridas/úlceras?", ["Não", "Sim"])
    s_feridas_esp = st.text_input("Especifique sobre as feridas:") if s_feridas == "Sim" else ""

    s_pelos = st.selectbox("Queda de pelos?", ["Não", "Sim"])
    s_frio = st.selectbox("Pés ou mãos frios?", ["Não", "Sim"])
    s_unhas = st.selectbox("Unhas fracas?", ["Não", "Sim"])

with tab5:
    st.subheader("Caracterização da Dor")
    d_tipo = st.multiselect("A dor apresenta:", ["Queimação", "Cansaço/Peso", "Cãibra", "Pontada", "Outro"])
    d_piora = st.text_input("A dor piora quando:")
    d_melhora = st.multiselect("A dor melhora quando:", ["Senta/descansa", "Eleva as pernas", "Fica em repouso/deitado", "Outro"])
    d_obs = st.text_area("Observações adicionais sobre a dor")

with tab6:
    st.subheader("Histórico Médico e Comorbidades")
    c_possui = st.selectbox("Possui comorbidades?", ["Não", "Sim"])
    c_desc = st.text_area("Descreva as comorbidades:") if c_possui == "Sim" else ""

    c_condicoes = st.multiselect("Condições presentes:", ["Hipertensão arterial", "Diabetes", "Dislipidemia", "Doenças do coração", "TVP", "AVC", "Doença renal", "Trombofilia", "Outra"])
    
    c_covid = st.selectbox("Teve COVID?", ["Não", "Sim"])
    c_vacina = st.selectbox("Vacinado para COVID?", ["Não", "Sim"])
    c_doses = st.text_input("Quantas doses?") if c_vacina == "Sim" else ""

with tab7:
    st.subheader("Cirurgias Prévias e Hábitos")
    cx_safena_retirada = st.selectbox("Retirada de safena", ["Não", "Sim"])
    cx_safena_ponte = st.selectbox("Pontes de safena", ["Não", "Sim"])
    cx_aplicacao_varizes = st.selectbox("Aplicação em varizes", ["Não", "Sim"])
    cx_cirurgia_varizes = st.selectbox("Cirurgia de varizes", ["Não", "Sim"])
    cx_arterial = st.selectbox("Cirurgia arterial", ["Não", "Sim"])
    cx_outras = st.text_input("Outras cirurgias")
    cx_medicamentos = st.text_area("Medicamentos contínuos")
    cx_hormonal = st.selectbox("Uso hormonal/anticoncepcional", ["Não se aplica", "Sim", "Não"])
    
    h_tabagismo = st.selectbox("Tabagismo", ["Nunca", "Ocasionalmente", "Diariamente"])
    h_tabagismo_detalhe = st.text_input("Detalhes do fumo (quantidade/tempo):") if h_tabagismo in ["Ocasionalmente", "Diariamente"] else ""

    h_atividade = st.selectbox("Atividade física", ["Não", "Sim"])
    h_atividade_freq = st.text_input("Quantas vezes por semana?") if h_atividade == "Sim" else ""

    h_alcool = st.selectbox("Consumo de álcool", ["Não", "Sim"])
    f_historico = st.multiselect("Histórico familiar:", ["Varizes graves", "Trombose", "Amputação", "Infarto/AVC precoce", "Aneurisma", "Nenhum"])

with tab8:
    st.subheader("Exames e Finalização")
    cl_exame = st.text_area("Exame clínico")
    cl_tratamento = st.text_area("Medicação/tratamento atual")
    
    ex_anexados = st.selectbox("Há exames anexados?", ["Não", "Sim"])
    ex_quais = st.text_input("Quais exames?") if ex_anexados == "Sim" else ""
    
    dg_hipotese = st.text_input("Hipótese diagnóstica")
    dg_duvida = st.text_area("Descreva objetivamente a dúvida clínica *")
    dg_obs = st.text_area("Observações complementares")
    
    imagens_exames = st.file_uploader("Selecione as fotos dos exames", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    st.markdown("---")
    if st.button("📄 Gerar Documento Word (.docx)", type="primary"):
        if not prof_nome.strip() or not pac_nome.strip() or not dg_duvida.strip():
            st.error("❌ Erro: Preencha os campos obrigatórios (*): Nome do Profissional, Nome do Paciente e Dúvida Clínica.")
        elif prof_cpf and not validar_cpf(prof_cpf):
            st.error("❌ Erro: O CPF do profissional é inválido.")
        elif pac_cpf and not validar_cpf(pac_cpf):
            st.error("❌ Erro: O CPF do paciente é inválido.")
        else:
            try:
                doc = Document()
                for section in doc.sections:
                    section.top_margin = Inches(1)
                    section.bottom_margin = Inches(1)
                    section.left_margin = Inches(1)
                    section.right_margin = Inches(1)

                logos_para_word = ["./fotos/logo1.png", "./fotos/logo2.png", "logo1.png", "logo2.png"]
                logo_principal = "./fotos/logo1.png" if os.path.exists("./fotos/logo1.png") else "logo1.png"
                
                # Aplica o cabeçalho e a marca d'água robusta de fundo
                adicionar_cabecalho_e_marca_dagua(doc, logo_principal, logos_para_word)

                p_title = doc.add_paragraph()
                p_title.paragraph_format.space_before = Pt(24)
                r_title = p_title.add_run("SOLICITAÇÃO DE TELEATENDIMENTO — TRIAGEM VASCULAR")
                r_title.bold = True
                r_title.font.size = Pt(14)
                r_title.font.color.rgb = RGBColor(23, 56, 50)

                p_sub = doc.add_paragraph()
                r_sub = p_sub.add_run("Telessaúde UEA")
                r_sub.italic = True
                r_sub.font.size = Pt(10)
                r_sub.font.color.rgb = RGBColor(44, 110, 99)

                doc.add_paragraph().paragraph_format.space_after = Pt(10)
                pac_idade = calcular_idade(pac_nasc)

                dados_gerais = [
                    ("PROFISSIONAL SOLICITANTE", [
                        ("Nome do profissional", prof_nome), ("CPF", prof_cpf), ("Profissão", prof_profissao),
                        ("Registro do conselho", prof_conselho), ("RQE", prof_rqe), ("Município", prof_municipio), ("E-mail", prof_email)
                    ]),
                    ("IDENTIFICAÇÃO DO PACIENTE", [
                        ("Nome do paciente", pac_nome), ("CPF do paciente", pac_cpf), ("Data de nascimento", pac_nasc),
                        ("Idade", pac_idade), ("Sexo", pac_sexo), ("Profissão / Ocupação", pac_ocupacao),
                        ("Peso", pac_peso), ("Altura", pac_altura), ("P.A.", pac_pa),
                        ("Paciente indígena", pac_indigena), ("Etnia", pac_etnia)
                    ]),
                    ("QUEIXA PRINCIPAL", [
                        ("Motivo da consulta", q_motivo), ("CID", q_cid), ("CIAP", q_ciap)
                    ]),
                    ("SINTOMAS VASCULARES", [
                        ("Em pé parado", s_empe), ("Sentado", s_sentado), ("Dor nas pernas", s_dorpernas),
                        ("Pernas incham", s_incham), ("Inchaço quando", s_inchaco_quando), 
                        ("Inchaço onde", s_inchaco_onde), ("Varizes", s_varizes), 
                        ("Pele escurecida", s_pele), ("Feridas", s_feridas), ("Especif. feridas", s_feridas_esp),
                        ("Queda de pelos", s_pelos), ("Pés/mãos frios", s_frio), ("Unhas fracas", s_unhas)
                    ]),
                    ("CARACTERIZAÇÃO DA DOR", [
                        ("Tipo de dor", ", ".join(d_tipo) if d_tipo else ""), ("Piora quando", d_piora),
                        ("Melhora quando", ", ".join(d_melhora) if d_melhora else ""), ("Observações da dor", d_obs)
                    ]),
                    ("HISTÓRICO MÉDICO E COMORBIDADES", [
                        ("Possui comorbidades", c_possui), ("Descrição comorbidades", c_desc),
                        ("Condições presentes", ", ".join(c_condicoes) if c_condicoes else ""), ("Teve COVID", c_covid),
                        ("Vacinado COVID", c_vacina), ("Doses", c_doses)
                    ]),
                    ("CIRURGIAS PRÉVIAS E MEDICAMENTOS", [
                        ("Retirada de safena", cx_safena_retirada), ("Pontes de safena", cx_safena_ponte),
                        ("Aplicação em varizes", cx_aplicacao_varizes), ("Cirurgia de varizes", cx_cirurgia_varizes),
                        ("Cirurgia arterial", cx_arterial), ("Outras cirurgias", cx_outras),
                        ("Medicamentos contínuos", cx_medicamentos), ("Hormonal", cx_hormonal)
                    ]),
                    ("HÁBITOS DE VIDA", [
                        ("Tabagismo", h_tabagismo), ("Detalhes fumo", h_tabagismo_detalhe),
                        ("Atividade física", h_atividade), ("Frequência atividade", h_atividade_freq), ("Álcool", h_alcool)
                    ]),
                    ("HISTÓRICO FAMILIAR", [
                        ("Histórico familiar", ", ".join(f_historico) if f_historico else "")
                    ]),
                    ("EXAME CLÍNICO E TRATAMENTO ATUAL", [
                        ("Exame clínico", cl_exame), ("Tratamento atual", cl_tratamento)
                    ]),
                    ("EXAMES E ANEXOS", [
                        ("Exames anexados", ex_anexados), ("Quais exames", ex_quais)
                    ]),
                    ("HIPÓTESE DIAGNÓSTICA E DÚVIDA CLÍNICA", [
                        ("Hipótese diagnóstica", dg_hipotese), ("Dúvida clínica", dg_duvida), ("Observações", dg_obs)
                    ])
                ]

                for sec_title, campos in dados_gerais:
                    p_sec = doc.add_paragraph()
                    r_sec = p_sec.add_run(sec_title)
                    r_sec.bold = True
                    r_sec.font.size = Pt(12)
                    r_sec.font.color.rgb = RGBColor(23, 56, 50)
                    p_sec.paragraph_format.space_before = Pt(12)
                    p_sec.paragraph_format.space_after = Pt(4)

                    for label, val in campos:
                        if val and str(val).strip() != "":
                            p_field = doc.add_paragraph()
                            p_field.paragraph_format.space_after = Pt(3)
                            r_label = p_field.add_run(f"{label}: ")
                            r_label.bold = True
                            r_label.font.size = Pt(10)
                            r_val = p_field.add_run(str(val))
                            r_val.font.size = Pt(10)

                if imagens_exames:
                    for idx, img_file in enumerate(imagens_exames):
                        doc.add_page_break()
                        p_desc = doc.add_paragraph()
                        r_d = p_desc.add_run(f"IMAGEM ANEXADA {idx + 1}")
                        r_d.bold = True
                        r_d.font.size = Pt(12)
                        r_d.font.color.rgb = RGBColor(23, 56, 50)
                        p_desc.paragraph_format.space_after = Pt(14)
                        try:
                            doc.add_picture(img_file, width=Inches(5.5))
                        except Exception:
                            pass

                file_name = f"Triagem_Vascular_{pac_nome.replace(' ', '_')}.docx"
                doc.save(file_name)
                
                st.success("✅ Documento Word gerado com sucesso!")
                with open(file_name, "rb") as f:
                    st.download_button(
                        label="📥 Baixar Documento Word (.docx)",
                        data=f,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )
            except Exception as e:
                st.error(f"❌ Erro ao gerar documento: {str(e)}")
