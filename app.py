import os
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from PIL import Image

# Configuração da página do Streamlit
st.set_page_config(
    page_title="Triagem Vascular · Telessaúde UEA",
    page_icon="🩺",
    layout="centered"
)

# Estilização visual limpa e profissional alinhada à identidade visual
st.markdown("""
    <style>
    .main-title { font-size: 26px; font-weight: bold; color: #173832; margin-bottom: 0px; }
    .sub-title { font-size: 14px; color: #4B615C; margin-bottom: 20px; }
    .section-header { font-size: 18px; font-weight: bold; color: #1F4F47; border-bottom: 2px solid #2C6E63; padding-bottom: 4px; margin-top: 20px; margin-bottom: 15px; }
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

# Cabeçalho com Logo Institucional (se houver logo.png na pasta)
col_logo, col_txt = st.columns([1, 4])
with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=110)
with col_txt:
    st.markdown('<p class="main-title">TELESSAÚDE UEA · Triagem Vascular</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Formulário de Solicitação de Teleatendimento</p>', unsafe_allow_html=True)

st.divider()

# Formulário único fluido por etapas visuais
with st.form("form_triagem"):
    
    # --- ETAPA 1 ---
    st.markdown('<p class="section-header">1. Profissional Solicitante</p>', unsafe_allow_html=True)
    prof_nome = st.text_input("Nome completo do profissional (não abreviar) *")
    col1, col2 = st.columns(2)
    with col1:
        prof_cpf = st.text_input("CPF")
        prof_conselho = st.text_input("Registro do conselho")
        prof_municipio = st.text_input("Município")
    with col2:
        prof_profissao = st.text_input("Profissão")
        prof_rqe = st.text_input("RQE (se médico)")
        prof_email = st.text_input("E-mail")

    # --- ETAPA 2 ---
    st.markdown('<p class="section-header">2. Identificação do Paciente</p>', unsafe_allow_html=True)
    pac_nome = st.text_input("Nome completo do paciente *")
    col3, col4 = st.columns(2)
    with col3:
        pac_cpf = st.text_input("CPF do paciente")
        pac_nasc = st.text_input("Data de nascimento (DD/MM/AAAA)")
        pac_sexo = st.selectbox("Sexo", ["", "Masculino", "Feminino"])
        pac_peso = st.text_input("Peso aproximado (kg)")
    with col4:
        pac_idade = st.text_input("Idade (calculada automaticamente ao preencher a data)")
        pac_ocupacao = st.text_input("Profissão / Ocupação")
        pac_altura = st.text_input("Altura (m)")
        pac_pa = st.text_input("P.A. (pressão arterial)")

    pac_indigena = st.selectbox("Paciente indígena?", ["Não", "Sim"])
    pac_etnia = ""
    if pac_indigena == "Sim":
        pac_etnia = st.text_input("Qual etnia?")

    # --- ETAPA 3 ---
    st.markdown('<p class="section-header">3. Queixa Principal</p>', unsafe_allow_html=True)
    q_motivo = st.text_area("Qual é o principal motivo da consulta? *")
    col5, col6 = st.columns(2)
    with col5:
        q_cid = st.text_input("CID")
    with col6:
        q_ciap = st.text_input("CIAP")

    # --- ETAPA 4 ---
    st.markdown('<p class="section-header">4. Sintomas Vasculares</p>', unsafe_allow_html=True)
    col7, col8 = st.columns(2)
    with col7:
        s_empe = st.selectbox("Permanece muito tempo em pé parado?", ["Não", "Sim"])
        s_dorpernas = st.selectbox("Sente dor nas pernas?", ["Não", "Sim"])
        s_varizes = st.selectbox("Nota veias dilatadas ou 'vasinhos' (varizes)?", ["Não", "Sim"])
        s_feridas = st.selectbox("Há feridas/úlceras nas pernas ou pés?", ["Não", "Sim"])
    with col8:
        s_sentado = st.selectbox("Permanece muito tempo sentado?", ["Não", "Sim"])
        s_incham = st.selectbox("As pernas incham?", ["Não", "Sim"])
        s_pele = st.selectbox("A pele das pernas ou tornozelos está escurecida/manchada?", ["Não", "Sim"])
        s_pelos = st.selectbox("Notou diminuição ou queda de pelos nas pernas/pés?", ["Não", "Sim"])

    s_inchaco_quando, s_inchaco_onde, s_feridas_esp = "", "", ""
    if s_incham == "Sim":
        s_inchaco_quando = st.selectbox("Se houver inchaço, ocorre mais:", ["", "Manhã", "Final do dia", "O tempo todo"])
        s_inchaco_onde = st.selectbox("O inchaço é em:", ["", "Apenas uma perna", "Ambas as pernas"])
    
    if s_feridas == "Sim":
        s_feridas_esp = st.text_input("Especifique sobre as feridas:")

    s_frio = st.selectbox("Sente os pés ou mãos muito frios em comparação ao resto do corpo?", ["Não", "Sim"])
    s_unhas = st.selectbox("Unhas dos pés fracas ou quebradiças?", ["Não", "Sim"])

    # --- ETAPA 5 ---
    st.markdown('<p class="section-header">5. Caracterização da Dor</p>', unsafe_allow_html=True)
    d_tipo = st.multiselect("A dor apresenta:", ["Queimação", "Cansaço/Peso", "Cãibra", "Pontada", "Outro"])
    d_tipo_outro = st.text_input("Especifique 'Outro' (Tipo de dor):") if "Outro" in d_tipo else ""
    
    d_piora = st.text_input("A dor piora quando:")
    
    d_melhora = st.multiselect("A dor melhora quando:", ["Senta/descansa", "Eleva as pernas", "Fica em repouso/deitado", "Outro"])
    d_melhora_outro = st.text_input("Especifique 'Outro' (Melhora da dor):") if "Outro" in d_melhora else ""
    
    d_obs = st.text_area("Observações adicionais sobre a dor")

    # --- ETAPA 6 ---
    st.markdown('<p class="section-header">6. Histórico Médico e Comorbidades</p>', unsafe_allow_html=True)
    c_possui = st.selectbox("Possui comorbidades?", ["Não", "Sim"])
    c_desc = st.text_area("Descreva as comorbidades:") if c_possui == "Sim" else ""
    
    c_condicoes = st.multiselect("Marque as condições presentes:", [
        "Hipertensão arterial", "Diabetes", "Colesterol/Triglicerídeos altos (Dislipidemia)", 
        "Doenças do coração", "Histórico de Trombose Venosa Profunda (TVP)", 
        "Histórico de AVC (derrame)", "Doença renal", "Trombofilia", "Outra"
    ])
    c_condicoes_outra = st.text_input("Especifique 'Outra' condição:") if "Outra" in c_condicoes else ""
    
    col9, col10 = st.columns(2)
    with col9:
        c_covid = st.selectbox("Teve COVID?", ["Não", "Sim"])
    with col10:
        c_vacina = st.selectbox("Vacinado para COVID?", ["Não", "Sim"])
    
    c_doses = st.text_input("Quantas doses?") if c_vacina == "Sim" else ""

    # --- ETAPA 7 ---
    st.markdown('<p class="section-header">7. Cirurgias Prévias e Medicamentos</p>', unsafe_allow_html=True)
    col11, col12 = st.columns(2)
    with col11:
        cx_safena_retirada = st.selectbox("Retirada de safena", ["Não", "Sim"])
        cx_aplicacao_varizes = st.selectbox("Aplicação em varizes", ["Não", "Sim"])
        cx_arterial = st.selectbox("Cirurgia arterial", ["Não", "Sim"])
    with col12:
        cx_safena_ponte = st.selectbox("Pontes de safena", ["Não", "Sim"])
        cx_cirurgia_varizes = st.selectbox("Cirurgia de varizes", ["Não", "Sim"])
    
    cx_outras = st.text_input("Outras cirurgias não listadas acima")
    cx_medicamentos = st.text_area("Quais medicamentos utiliza continuamente?")
    cx_hormonal = st.selectbox("Para mulheres: uso de anticoncepcional ou reposição hormonal?", ["Não se aplica", "Sim", "Não"])

    # --- ETAPA 8 ---
    st.markdown('<p class="section-header">8. Hábitos de Vida</p>', unsafe_allow_html=True)
    h_tabagismo = st.selectbox("Tabagismo / fumo", ["Nunca", "Ocasionalmente", "Diariamente"])
    h_tabagismo_detalhe = st.text_input("Detalhes do fumo (quantidade/tempo):") if h_tabagismo in ["Ocasionalmente", "Diariamente"] else ""
    
    col13, col14 = st.columns(2)
    with col13:
        h_atividade = st.selectbox("Pratica atividade física?", ["Não", "Sim"])
        h_atividade_freq = st.text_input("Quantas vezes por semana?") if h_atividade == "Sim" else ""
    with col14:
        h_alcool = st.selectbox("Consumo de álcool", ["Não", "Sim"])

    # --- ETAPA 9 ---
    st.markdown('<p class="section-header">9. Histórico Familiar</p>', unsafe_allow_html=True)
    f_historico = st.multiselect("Histórico familiar de:", [
        "Varizes graves", "Trombose", "Amputação por problema circulatório", 
        "Infarto ou AVC antes dos 60 anos", "Aneurisma", "Nenhum dos citados"
    ])

    # --- ETAPA 10 ---
    st.markdown('<p class="section-header">10. Exame Clínico e Tratamento Atual</p>', unsafe_allow_html=True)
    cl_exame = st.text_area("Exame clínico")
    cl_tratamento = st.text_area("Medicação / tratamento atual")

    # --- ETAPA 11 ---
    st.markdown('<p class="section-header">11. Exames e Anexos (Fotos)</p>', unsafe_allow_html=True)
    ex_anexados = st.selectbox("Há exames anexados?", ["Não", "Sim"])
    ex_quais = st.text_area("Quais exames estão em anexo?") if ex_anexados == "Sim" else ""
    
    uploaded_images = st.file_uploader("Selecione as fotos dos exames (cada foto ficará em uma página separada no Word):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    # --- ETAPA 12 ---
    st.markdown('<p class="section-header">12. Hipótese Diagnóstica e Dúvida Clínica</p>', unsafe_allow_html=True)
    dg_hipotese = st.text_area("Hipótese diagnóstica")
    dg_duvida = st.text_area("Descreva objetivamente a dúvida clínica *")
    dg_obs = st.text_area("Observações complementares")

    # Botão de Envio do Formulário
    submitted = st.form_submit_button("Gerar Documento Word (.docx)")

# Processamento ao submeter
if submitted:
    # Validações obrigatórias
    if not prof_nome.strip() or not pac_nome.strip() or not dg_duvida.strip():
        st.error("Por favor, preencha os campos obrigatórios marcados com asterisco (*): Nome do Profissional, Nome do Paciente e Dúvida Clínica.")
    elif prof_cpf and not validar_cpf(prof_cpf):
        st.error("O CPF do profissional informado é inválido.")
    elif pac_cpf and not validar_cpf(pac_cpf):
        st.error("O CPF do paciente informado é inválido.")
    else:
        try:
            doc = Document()
            for section in doc.sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)

            # Logo no documento Word se existir
            if os.path.exists("logo.png"):
                try:
                    doc.add_picture("logo.png", width=Inches(1.8))
                except Exception:
                    pass

            p_title = doc.add_paragraph()
            r_title = p_title.add_run("SOLICITAÇÃO DE TELEATENDIMENTO — TRIAGEM VASCULAR")
            r_title.bold = True
            r_title.font.size = Pt(16)
            r_title.font.color.rgb = RGBColor(23, 56, 50)

            p_sub = doc.add_paragraph()
            r_sub = p_sub.add_run("Telessaúde UEA")
            r_sub.italic = True
            r_sub.font.size = Pt(11)
            r_sub.font.color.rgb = RGBColor(44, 110, 99)

            doc.add_paragraph().paragraph_format.space_after = Pt(10)

            # Mapeamento dos dados coletados
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
                    ("Pernas incham", s_incham), ("Inchaço quando", s_inchaco_quando), ("Inchaço onde", s_inchaco_onde),
                    ("Varizes", s_varizes), ("Pele escurecida", s_pele), ("Feridas", s_feridas), ("Especif. feridas", s_feridas_esp),
                    ("Queda de pelos", s_pelos), ("Pés/mãos frios", s_frio), ("Unhas fracas", s_unhas)
                ]),
                ("CARACTERIZAÇÃO DA DOR", [
                    ("Tipo de dor", ", ".join(d_tipo) if isinstance(d_tipo, list) else d_tipo),
                    ("Especif. Outro (Dor)", d_tipo_outro), ("Piora quando", d_piora),
                    ("Melhora quando", ", ".join(d_melhora) if isinstance(d_melhora, list) else d_melhora),
                    ("Especif. Outro (Melhora)", d_melhora_outro), ("Observações da dor", d_obs)
                ]),
                ("HISTÓRICO MÉDICO E COMORBIDADES", [
                    ("Possui comorbidades", c_possui), ("Descrição comorbidades", c_desc),
                    ("Condições presentes", ", ".join(c_condicoes) if isinstance(c_condicoes, list) else c_condicoes),
                    ("Especif. Outra condição", c_condicoes_outra), ("Teve COVID", c_covid),
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
                    ("Histórico familiar", ", ".join(f_historico) if isinstance(f_historico, list) else f_historico)
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
                r_sec.font.size = Pt(13)
                r_sec.font.color.rgb = RGBColor(23, 56, 50)
                p_sec.paragraph_format.space_before = Pt(14)
                p_sec.paragraph_format.space_after = Pt(4)

                has_val = False
                for label, val in campos:
                    if val and str(val).strip() != "":
                        has_val = True
                        p_field = doc.add_paragraph()
                        p_field.paragraph_format.space_after = Pt(3)
                        r_label = p_field.add_run(f"{label}: ")
                        r_label.bold = True
                        r_label.font.size = Pt(11)
                        r_val = p_field.add_run(str(val))
                        r_val.font.size = Pt(11)

            # Inserção das Imagens (Cada uma isolada em sua própria página)
            if uploaded_images:
                for idx, img_file in enumerate(uploaded_images):
                    doc.add_page_break()
                    p_desc = doc.add_paragraph()
                    r_d = p_desc.add_run(f"IMAGEM ANEXADA {idx + 1}: {img_file.name}")
                    r_d.bold = True
                    r_d.font.size = Pt(13)
                    r_d.font.color.rgb = RGBColor(23, 56, 50)
                    p_desc.paragraph_format.space_after = Pt(14)

                    try:
                        doc.add_picture(img_file, width=Inches(5.5))
                    except Exception as img_err:
                        p_err = doc.add_paragraph()
                        p_err.add_run(f"[Erro ao carregar a imagem: {img_err}]").font.color.rgb = RGBColor(181, 80, 46)

            file_name = f"Triagem_Vascular_{pac_nome.replace(' ', '_')}.docx"
            doc.save(file_name)

            st.success("Documento Word gerado com sucesso!")
            
            # Botão de Download direto para o navegador do usuário
            with open(file_name, "rb") as file:
                st.download_button(
                    label="📥 Baixar Documento Word Agora",
                    data=file,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"Ocorreu um erro ao gerar o documento: {str(e)}")