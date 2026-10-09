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

# --- CSS CUSTOMIZADO PARA ACESSIBILIDADE E PROFISSIONAIS DE SAÚDE ---
st.markdown("""
    <style>
        .main {
            background-color: #F4F6F8;
        }
        html, body, [class*="css"] {
            font-size: 1.1rem !important;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            color: #1A252C;
        }
        h1 {
            color: #133832 !important;
            font-size: 2.2rem !important;
            font-weight: 700 !important;
        }
        h2, h3 {
            color: #133832 !important;
            font-weight: 600 !important;
        }
        .stTextInput input, .stTextArea textarea, .stSelectbox select {
            font-size: 1.15rem !important;
            padding: 12px !important;
            border-radius: 8px !important;
            border: 2px solid #BDC3C7 !important;
        }
        .stTextInput input:focus, .stTextArea textarea:focus {
            border-color: #133832 !important;
            box-shadow: 0 0 8px rgba(19, 56, 50, 0.2) !important;
        }
        label div p {
            font-size: 1.1rem !important;
            font-weight: 600 !important;
            color: #2C3E50 !important;
        }
        .stButton>button {
            background-color: #133832 !important;
            color: white !important;
            border-radius: 10px !important;
            font-size: 1.3rem !important;
            font-weight: 700 !important;
            padding: 0.8rem 2rem !important;
            width: 100%;
            box-shadow: 0 4px 12px rgba(19,56,50,0.3);
        }
        .stButton>button:hover {
            background-color: #1D544C !important;
        }
        /* Efeito de destaque para instruções importantes */
        .aviso-instrucao {
            background-color: #E8F4F8;
            border-left: 6px solid #133832;
            padding: 18px;
            border-radius: 8px;
            margin-top: 20px;
            margin-bottom: 20px;
            font-size: 1.15rem;
            color: #133832;
            font-weight: 600;
        }
    </style>
""", unsafe_allow_html=True)

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

def adicionar_cabecalho_e_marca_dagua(doc, logo_marca_dagua):
    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0]
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    if logo_marca_dagua and os.path.exists(logo_marca_dagua):
        try:
            run = header_para.add_run()
            run.add_picture(logo_marca_dagua, width=Inches(5.0))
        except Exception:
            pass

# --- CABEÇALHO VISUAL NO STREAMLIT (APENAS UMA LOGO) ---
img1_b64 = img_to_base64("./fotos/Design sem nome (12).png") or img_to_base64("Design sem nome (12).png")

logos_html = '<div style="display: flex; justify-content: center; align-items: center; gap: 15px; margin-bottom: 10px;">'
if img1_b64:
    logos_html += f'<img src="data:image/png;base64,{img1_b64}" style="height: 95px; width: auto;" />'
logos_html += '</div>'

st.markdown(logos_html, unsafe_allow_html=True)
st.markdown("<h1 style='text-align: center;'>Telessaúde UEA · Triagem Vascular</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.25rem; color: #4A5568;'>Preencha as informações abaixo passo a passo para gerar a solicitação oficial de teleatendimento.</p>", unsafe_allow_html=True)
st.markdown("---")

# --- SIDEBAR DE APOIO COM ORIENTAÇÕES CLARAS ---
with st.sidebar:
    st.image("https://img.icons8.com/color/96/stethoscope.png", width=70)
    st.markdown("### 📋 Orientações de Uso")
    st.success("Preencha os campos com atenção. Os itens marcados com asterisco (**\***) são de preenchimento obrigatório.")
    st.markdown("---")
    st.markdown("💡 **Dica:** Role a página de cima para baixo preenchendo as seções. Ao final, clique no botão verde para gerar e baixar o documento pronto.")
    st.markdown("---")
    st.caption("Suporte Técnico — Telessaúde UEA Contato: (92) 99209-7534   \n E-mail: telessaude@uea.edu.br" )


# ==========================================
# SEÇÃO 1: PROFISSIONAL SOLICITANTE
# ==========================================
st.markdown("### 1️⃣ Dados do Profissional Solicitante")
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        prof_nome = st.text_input("Nome completo do profissional (não abreviar) *", placeholder="Ex: Dr. João da Silva")
        prof_cpf = st.text_input("CPF do Profissional", placeholder="000.000.000-00")
        prof_profissao = st.text_input("Profissão", placeholder="Ex: Médico(a) / Enfermeiro(a)")
        prof_conselho = st.text_input("Registro do conselho", placeholder="Ex: CRM/AM 0000")
    with col2:
        prof_rqe = st.text_input("RQE (se médico)", placeholder="Ex: 00000")
        prof_municipio = st.text_input("Município de atuação", placeholder="Ex: Manaus")
        prof_email = st.text_input("E-mail profissional", placeholder="exemplo@email.com")

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# SEÇÃO 2: IDENTIFICAÇÃO DO PACIENTE
# ==========================================
st.markdown("### 2️⃣ Identificação do Paciente")
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        pac_nome = st.text_input("Nome completo do paciente *", placeholder="Nome completo sem abreviações")
        pac_cpf = st.text_input("CPF do paciente", placeholder="000.000.000-00")
        pac_nasc = st.text_input("Data de nascimento (DD/MM/AAAA) *", placeholder="DD/MM/AAAA")
        pac_sexo = st.selectbox("Sexo", ["", "Masculino", "Feminino"])
        pac_ocupacao = st.text_input("Profissão / Ocupação do paciente")
    with col2:
        c_peso, c_alt, c_pa = st.columns(3)
        with c_peso:
            pac_peso = st.text_input("Peso (kg)", placeholder="70")
        with c_alt:
            pac_altura = st.text_input("Altura (m)", placeholder="1.70")
        with c_pa:
            pac_pa = st.text_input("P.A.", placeholder="120/80")
            
        pac_indigena = st.selectbox("Paciente indígena?", ["Não", "Sim"])
        pac_etnia = st.text_input("Qual etnia?") if pac_indigena == "Sim" else ""

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# SEÇÃO 3: QUEIXA PRINCIPAL
# ==========================================
st.markdown("### 3️⃣ Queixa Principal")
with st.container():
    q_motivo = st.text_area("Qual é o principal motivo da consulta? *", placeholder="Descreva detalhadamente o quadro atual e a queixa principal...", height=130)
    col1, col2 = st.columns(2)
    with col1:
        q_cid = st.text_input("CID (se houver)", placeholder="Ex: I83.9")
    with col2:
        q_ciap = st.text_input("CIAP (se houver)", placeholder="Ex: K93")

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# SEÇÃO 4: SINTOMAS VASCULARES
# ==========================================
st.markdown("### 4️⃣ Sintomas Vasculares")
with st.container():
    col1, col2, col3 = st.columns(3)
    with col1:
        s_empe = st.selectbox("Permanece muito tempo em pé parado?", ["Não", "Sim"])
        s_sentado = st.selectbox("Permanece muito tempo sentado?", ["Não", "Sim"])
        s_dorpernas = st.selectbox("Sente dor nas pernas?", ["Não", "Sim"])
    with col2:
        s_incham = st.selectbox("As pernas incham?", ["Não", "Sim"])
        s_varizes = st.selectbox("Nota veias dilatadas ou varizes?", ["Não", "Sim"])
        s_pele = st.selectbox("Pele escurecida/manchada?", ["Não", "Sim"])
    with col3:
        s_pelos = st.selectbox("Queda de pelos?", ["Não", "Sim"])
        s_frio = st.selectbox("Pés ou mãos frios?", ["Não", "Sim"])
        s_unhas = st.selectbox("Unhas fracas?", ["Não", "Sim"])

    if s_incham == "Sim":
        st.info("ℹ️ **Detalhes do Inchaço:**")
        ic_1, ic_2 = st.columns(2)
        with ic_1:
            s_inchaco_quando = st.selectbox("Inchaço ocorre mais:", ["Manhã", "Final do dia", "O tempo todo"])
        with ic_2:
            s_inchaco_onde = st.selectbox("Inchaço é em:", ["Apenas uma perna", "Ambas as pernas"])
    else:
        s_inchaco_quando, s_inchaco_onde = "", ""

    s_feridas = st.selectbox("Há feridas/úlceras?", ["Não", "Sim"])
    s_feridas_esp = st.text_input("Especifique sobre as feridas:") if s_feridas == "Sim" else ""

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# SEÇÃO 5: CARACTERIZAÇÃO DA DOR
# ==========================================
st.markdown("### 5️⃣ Caracterização da Dor")
with st.container():
    d_tipo = st.multiselect("A dor apresenta:", ["Queimação", "Cansaço/Peso", "Cãibra", "Pontada", "Outro"])
    d_piora = st.text_input("A dor piora quando:", placeholder="Ex: Ao caminhar longas distâncias, ao ficar muito tempo em pé...")
    d_melhora = st.multiselect("A dor melhora quando:", ["Senta/descansa", "Eleva as pernas", "Fica em repouso/deitado", "Outro"])
    d_obs = st.text_area("Observações adicionais sobre a dor", placeholder="Detalhes complementares sobre a dor...")

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# SEÇÃO 6: HISTÓRICO MÉDICO E COMORBIDADES
# ==========================================
st.markdown("### 6️⃣ Histórico Médico e Comorbidades")
with st.container():
    c_possui = st.selectbox("Possui comorbidades?", ["Não", "Sim"])
    c_desc = st.text_area("Descreva as comorbidades:") if c_possui == "Sim" else ""

    c_condicoes = st.multiselect("Condições presentes:", ["Hipertensão arterial", "Diabetes", "Dislipidemia", "Doenças do coração", "TVP", "AVC", "Doença renal", "Trombofilia", "Outra"])
    
    col1, col2 = st.columns(2)
    with col1:
        c_covid = st.selectbox("Teve COVID?", ["Não", "Sim"])
    with col2:
        c_vacina = st.selectbox("Vacinado para COVID?", ["Não", "Sim"])
    
    c_doses = st.text_input("Quantas doses tomou?") if c_vacina == "Sim" else ""

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# SEÇÃO 7: CIRURGIAS E HÁBITOS DE VIDA
# ==========================================
st.markdown("### 7️⃣ Cirurgias Prévias e Hábitos de Vida")
with st.container():
    st.markdown("#### 🔪 Intervenções Cirúrgicas")
    cx_1, cx_2, cx_3 = st.columns(3)
    with cx_1:
        cx_safena_retirada = st.selectbox("Retirada de safena", ["Não", "Sim"])
        cx_safena_ponte = st.selectbox("Pontes de safena", ["Não", "Sim"])
    with cx_2:
        cx_aplicacao_varizes = st.selectbox("Aplicação em varizes", ["Não", "Sim"])
        cx_cirurgia_varizes = st.selectbox("Cirurgia de varizes", ["Não", "Sim"])
    with cx_3:
        cx_arterial = st.selectbox("Cirurgia arterial", ["Não", "Sim"])
        cx_hormonal = st.selectbox("Uso hormonal/anticoncepcional", ["Não se aplica", "Sim", "Não"])

    cx_outras = st.text_input("Outras cirurgias relevantes")
    cx_medicamentos = st.text_area("Medicamentos em uso contínuo", placeholder="Liste os medicamentos que o paciente faz uso...")

    st.markdown("#### 🏃 Hábitos e Histórico Familiar")
    hb_1, hb_2, hb_3 = st.columns(3)
    with hb_1:
        h_tabagismo = st.selectbox("Tabagismo", ["Nunca", "Ocasionalmente", "Diariamente"])
    with hb_2:
        h_atividade = st.selectbox("Atividade física", ["Não", "Sim"])
    with hb_3:
        h_alcool = st.selectbox("Consumo de álcool", ["Não", "Sim"])

    h_tabagismo_detalhe = st.text_input("Detalhes do fumo (quantidade/tempo):") if h_tabagismo in ["Ocasionalmente", "Diariamente"] else ""
    h_atividade_freq = st.text_input("Quantas vezes por semana faz atividade?") if h_atividade == "Sim" else ""

    f_historico = st.multiselect("Histórico familiar:", ["Varizes graves", "Trombose", "Amputação", "Infarto/AVC precoce", "Aneurisma", "Nenhum"])

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# SEÇÃO 8: EXAMES E FINALIZAÇÃO
# ==========================================
st.markdown("### 8️⃣ Exames, Dúvida Clínica e Finalização")
with st.container():
    cl_exame = st.text_area("Achados do Exame Clínico físico")
    cl_tratamento = st.text_area("Medicação ou tratamento atual em curso")
    
    col1, col2 = st.columns(2)
    with col1:
        ex_anexados = st.selectbox("Há exames complementares anexados?", ["Não", "Sim"])
    with col2:
        ex_quais = st.text_input("Quais exames?") if ex_anexados == "Sim" else ""
    
    dg_hipotese = st.text_input("Hipótese diagnóstica", placeholder="Ex: Insuficiência venosa crônica")
    dg_duvida = st.text_area("Descreva objetivamente a dúvida clínica para o teleatendimento *", placeholder="Qual a principal dúvida para o especialista vascular?", height=130)
    dg_obs = st.text_area("Observações complementares gerais")
    
    imagens_exames = st.file_uploader("Selecione as fotos dos exames para anexar ao documento", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    st.markdown("---")
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Seta indicativa chamativa apontando para o botão de gerar
    st.markdown("<p style='text-align: center; font-size: 1.5rem; font-weight: bold; color: #133832;'>👇 CLIQUE NO BOTÃO ABAIXO PARA GERAR O DOCUMENTO 👇</p>", unsafe_allow_html=True)
    
    gerar_doc = st.button("📄 GERAR DOCUMENTO WORD OFICIAL (.DOCX)", type="primary")

    if gerar_doc:
        if not prof_nome.strip() or not pac_nome.strip() or not dg_duvida.strip():
            st.error("❌ **Atenção:** Preencha os campos obrigatórios (*): Nome do Profissional, Nome do Paciente e Dúvida Clínica.")
        elif prof_cpf and not validar_cpf(prof_cpf):
            st.error("❌ **Atenção:** O CPF do profissional informado é inválido.")
        elif pac_cpf and not validar_cpf(pac_cpf):
            st.error("❌ **Atenção:** O CPF do paciente informado é inválido.")
        else:
            try:
                with st.spinner("Gerando documento formatado e processando imagens... Aguarde um instante."):
                    doc = Document()
                    for section in doc.sections:
                        section.top_margin = Inches(1)
                        section.bottom_margin = Inches(1)
                        section.left_margin = Inches(1)
                        section.right_margin = Inches(1)

                    logo_path = "./fotos/Design sem nome (12).png" if os.path.exists("./fotos/Design sem nome (12).png") else "Design sem nome (12).png"
                    adicionar_cabecalho_e_marca_dagua(doc, logo_path)

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
                            ("Pernas incham", s_incham), ("Inchaço когда", s_inchaco_quando), 
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
                    
                    st.success("✅ **Documento Word gerado com sucesso!**")
                    
                    # Seta chamativa apontando para o botão de download
                    st.markdown("<p style='text-align: center; font-size: 1.6rem; font-weight: bold; color: #D9534F; margin-top: 15px;'>👇 ⬇️ CLIQUE NO BOTÃO ABAIXO PARA BAIXAR O SEU DOCUMENTO ⬇️ 👇</p>", unsafe_allow_html=True)
                    
                    with open(file_name, "rb") as f:
                        st.download_button(
                            label="📥 BAIXAR DOCUMENTO WORD (.DOCX)",
                            data=f,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    
                    # Instrução exata solicitada logo após a geração do documento
                    st.markdown("""
                        <div class="aviso-instrucao">
                            ⚠️ <strong>INSTRUÇÃO IMPORTANTE:</strong><br>
                            Após baixar o documento preenchido, encaminhe-o junto com os anexos (fotos, exames laboratoriais, raio-x e outros documentos que sejam necessários) diretamente para o e-mail: 
                            <a href="mailto:telessaude@uea.edu.br" style="color: #0E2923; text-decoration: underline;">telessaude@uea.edu.br</a>
                        </div>
                    """, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"❌ Erro ao gerar documento: {str(e)}")
