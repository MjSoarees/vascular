import os
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from PIL import Image

# Configuração global de aparência
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("green")

def validar_cpf(cpf):
    cpf = ''.join(filter(str.isdigit, cpf))
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

class TriagemVascularApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Triagem Vascular · Telessaúde UEA")
        self.geometry("1120x720")
        self.configure(fg_color="#F5F7F6")

        self.current_step = 0
        self.form_data = {}
        self.selected_images = []
        self.widgets_map = {}
        self.conditional_frames = {}
        self.saved_file_path = ""

        self.steps = [
            ("Profissional solicitante", "Dados de quem está solicitando o teleatendimento.", [
                ("prof_nome", "Nome completo do profissional (não abreviar)", "entry", [], True, False),
                ("prof_cpf", "CPF", "cpf", [], False, False),
                ("prof_profissao", "Profissão", "entry", [], False, False),
                ("prof_conselho", "Registro do conselho", "entry", [], False, False),
                ("prof_rqe", "RQE (se médico)", "entry", [], False, False),
                ("prof_municipio", "Município", "entry", [], False, False),
                ("prof_email", "E-mail", "entry", [], False, False),
            ]),
            ("Identificação do paciente", "Dados do paciente encaminhado para avaliação vascular.", [
                ("pac_nome", "Nome completo", "entry", [], True, False),
                ("pac_cpf", "CPF", "cpf", [], False, False),
                ("pac_nasc", "Data de nascimento", "date", [], False, False),
                ("pac_idade", "Idade", "entry", [], False, False),
                ("pac_sexo", "Sexo", "pills", ["Masculino", "Feminino"], False, False),
                ("pac_ocupacao", "Profissão / Ocupação", "entry", [], False, False),
                ("pac_peso", "Peso aproximado (kg)", "entry", [], False, False),
                ("pac_altura", "Altura (m)", "entry", [], False, False),
                ("pac_pa", "P.A. (pressão arterial)", "entry", [], False, False),
                ("pac_indigena", "Paciente indígena?", "pills", ["Sim", "Não"], False, False),
                ("pac_etnia", "Qual etnia?", "entry", [], False, False, "pac_indigena", "Sim"),
            ]),
            ("Queixa principal", "História da doença atual, de preferência nas palavras do paciente.", [
                ("q_motivo", "Qual é o principal motivo da consulta?", "text", [], True, False),
                ("q_cid", "CID", "entry", [], False, False),
                ("q_ciap", "CIAP", "entry", [], False, False),
            ]),
            ("Sintomas vasculares", "Marque o que se aplica ao quadro do paciente.", [
                ("s_empe", "Permanece muito tempo em pé parado?", "pills", ["Sim", "Não"], False, False),
                ("s_sentado", "Permanece muito tempo sentado?", "pills", ["Sim", "Não"], False, False),
                ("s_dorpernas", "Sente dor nas pernas?", "pills", ["Sim", "Não"], False, False),
                ("s_incham", "As pernas incham?", "pills", ["Sim", "Não"], False, False),
                ("s_inchaco_quando", "Se houver inchaço, ocorre mais:", "pills", ["Manhã", "Final do dia", "O tempo todo"], False, False, "s_incham", "Sim"),
                ("s_inchaco_onde", "O inchaço é em:", "pills", ["Apenas uma perna", "Ambas as pernas"], False, False, "s_incham", "Sim"),
                ("s_varizes", "Nota veias dilatadas ou 'vasinhos' (varizes)?", "pills", ["Sim", "Não"], False, False),
                ("s_pele", "A pele das pernas ou tornozelos está escurecida/manchada?", "pills", ["Sim", "Não"], False, False),
                ("s_feridas", "Há feridas/úlceras nas pernas ou pés?", "pills", ["Sim", "Não"], False, False),
                ("s_feridas_esp", "Especifique sobre as feridas:", "text", [], False, False, "s_feridas", "Sim"),
                ("s_pelos", "Notou diminuição ou queda de pelos nas pernas/pés?", "pills", ["Sim", "Não"], False, False),
                ("s_frio", "Sente os pés ou mãos muito frios em comparação ao resto do corpo?", "pills", ["Sim", "Não"], False, False),
                ("s_unhas", "Unhas dos pés fracas ou quebradiças?", "pills", ["Sim", "Não"], False, False),
            ]),
            ("Caracterização da dor", "Detalhes sobre o padrão de dor relatado.", [
                ("d_tipo", "A dor apresenta:", "multipills", ["Queimação", "Cansaço/Peso", "Cãibra", "Pontada", "Outro"], False, False),
                ("d_tipo_outro", "Especifique 'Outro' (Tipo de dor):", "entry", [], False, False, "d_tipo", "Outro"),
                ("d_piora", "A dor piora quando:", "text", [], False, False),
                ("d_melhora", "A dor melhora quando:", "multipills", ["Senta/descansa", "Eleva as pernas", "Fica em repouso/deitado", "Outro"], False, False),
                ("d_melhora_outro", "Especifique 'Outro' (Melhora da dor):", "entry", [], False, False, "d_melhora", "Outro"),
                ("d_obs", "Observações adicionais sobre a dor", "text", [], False, False),
            ]),
            ("Histórico médico e comorbidades", "Condições clínicas e fatores de risco relevantes.", [
                ("c_possui", "Possui comorbidades?", "pills", ["Sim", "Não"], False, False),
                ("c_desc", "Descreva as comorbidades:", "text", [], False, False, "c_possui", "Sim"),
                ("c_condicoes", "Marque as condições presentes:", "multipills", [
                    "Hipertensão arterial", "Diabetes", "Colesterol/Triglicerídeos altos (Dislipidemia)", 
                    "Doenças do coração", "Histórico de Trombose Venosa Profunda (TVP)", 
                    "Histórico de AVC (derrame)", "Doença renal", "Trombofilia", "Outra"
                ], False, False),
                ("c_condicoes_outra", "Especifique 'Outra' condição:", "entry", [], False, False, "c_condicoes", "Outra"),
                ("c_covid", "Teve COVID?", "pills", ["Sim", "Não"], False, False),
                ("c_vacina", "Vacinado para COVID?", "pills", ["Sim", "Não"], False, False),
                ("c_doses", "Quantas doses?", "entry", [], False, False, "c_vacina", "Sim"),
            ]),
            ("Cirurgias prévias e medicamentos", "Procedimentos vasculares anteriores e uso atual de medicamentos.", [
                ("cx_safena_retirada", "Retirada de safena", "pills", ["Sim", "Não"], False, False),
                ("cx_safena_ponte", "Pontes de safena", "pills", ["Sim", "Não"], False, False),
                ("cx_aplicacao_varizes", "Aplicação em varizes", "pills", ["Sim", "Não"], False, False),
                ("cx_cirurgia_varizes", "Cirurgia de varizes", "pills", ["Sim", "Não"], False, False),
                ("cx_arterial", "Cirurgia arterial", "pills", ["Sim", "Não"], False, False),
                ("cx_outras", "Outras cirurgias não listadas acima", "text", [], False, False),
                ("cx_medicamentos", "Quais medicamentos utiliza continuamente?", "text", [], False, False),
                ("cx_hormonal", "Para mulheres: uso de anticoncepcional ou reposição hormonal?", "pills", ["Sim", "Não", "Não se aplica"], False, False),
            ]),
            ("Hábitos de vida", "Tabagismo, atividade física e consumo de álcool.", [
                ("h_tabagismo", "Tabagismo / fumo", "pills", ["Nunca", "Ocasionalmente", "Diariamente"], False, False),
                ("h_tabagismo_detalhe", "Detalhes do fumo (quantidade/tempo):", "text", [], False, False, "h_tabagismo", ["Ocasionalmente", "Diariamente"]),
                ("h_atividade", "Pratica atividade física?", "pills", ["Sim", "Não"], False, False),
                ("h_atividade_freq", "Quantas vezes por semana?", "entry", [], False, False, "h_atividade", "Sim"),
                ("h_alcool", "Consumo de álcool", "pills", ["Sim", "Não"], False, False),
            ]),
            ("Histórico familiar", "Antecedentes em pais, irmãos ou avós.", [
                ("f_historico", "Histórico familiar de:", "multipills", [
                    "Varizes graves", "Trombose", "Amputação por problema circulatório", 
                    "Infarto ou AVC antes dos 60 anos", "Aneurisma", "Nenhum dos citados"
                ], False, False),
            ]),
            ("Exame clínico e tratamento atual", "Achados do exame físico e conduta em curso.", [
                ("cl_exame", "Exame clínico", "text", [], False, False),
                ("cl_tratamento", "Medicação / tratamento atual", "text", [], False, False),
            ]),
            ("Exames e anexos", "Documentos complementares e fotos para o documento Word.", [
                ("ex_anexados", "Há exames anexados?", "pills", ["Sim", "Não"], False, False),
                ("ex_quais", "Quais exames estão em anexo?", "text", [], False, False, "ex_anexados", "Sim"),
            ]),
            ("Hipótese diagnóstica e dúvida", "O núcleo da solicitação: o que precisa ser respondido.", [
                ("dg_hipotese", "Hipótese diagnóstica", "text", [], False, False),
                ("dg_duvida", "Descreva objetivamente a dúvida clínica", "text", [], True, True),
                ("dg_obs", "Observações complementares", "text", [], False, False),
            ])
        ]

        self.setup_layout()

    def setup_layout(self):
        self.rail = ctk.CTkFrame(self, width=280, fg_color="#173832", corner_radius=0)
        self.rail.pack(side="left", fill="y")
        self.rail.pack_propagate(False)

        # Inserção da Logo na Barra Lateral (se o arquivo logo.png existir na pasta)
        logo_path = "logo.png"
        if os.path.exists(logo_path):
            try:
                img_pil = Image.open(logo_path)
                self.logo_img = ctk.CTkImage(light_image=img_pil, dark_image=img_pil, size=(160, 50))
                logo_lbl = ctk.CTkLabel(self.rail, image=self.logo_img, text="")
                logo_lbl.pack(anchor="w", padx=24, pady=(28, 4))
            except Exception:
                pass

        brand_lbl = ctk.CTkLabel(self.rail, text="TELESSAÚDE · UEA", font=("Courier", 11, "bold"), text_color="#89B3A8")
        brand_lbl.pack(anchor="w", padx=24, pady=(16, 4))

        title_lbl = ctk.CTkLabel(self.rail, text="Triagem Vascular", font=("Georgia", 20, "bold"), text_color="#F4F8F6")
        title_lbl.pack(anchor="w", padx=24, pady=(0, 2))

        sub_lbl = ctk.CTkLabel(self.rail, text="Solicitação de teleatendimento", font=("Arial", 12), text_color="#9EBBB2")
        sub_lbl.pack(anchor="w", padx=24, pady=(0, 24))

        self.steps_frame = ctk.CTkScrollableFrame(self.rail, fg_color="transparent", scrollbar_button_color="#1F4F47")
        self.steps_frame.pack(fill="both", expand=True, padx=10, pady=(0, 20))

        self.step_buttons = []
        for i, (title, _, _) in enumerate(self.steps):
            btn = ctk.CTkButton(
                self.steps_frame,
                text=f"{str(i+1).zfill(2)}  {title}",
                font=("Arial", 13),
                anchor="w",
                fg_color="transparent",
                text_color="#9EBBB2",
                hover_color="#1E463E",
                command=lambda idx=i: self.jump_to_step(idx)
            )
            btn.pack(fill="x", pady=2)
            self.step_buttons.append(btn)

        self.review_btn = ctk.CTkButton(
            self.steps_frame,
            text=f"{str(len(self.steps)+1).zfill(2)}  Revisão e Geração",
            font=("Arial", 13),
            anchor="w",
            fg_color="transparent",
            text_color="#9EBBB2",
            hover_color="#1E463E",
            command=lambda: self.jump_to_step(len(self.steps))
        )
        self.review_btn.pack(fill="x", pady=2)

        self.main_area = ctk.CTkFrame(self, fg_color="#F5F7F6", corner_radius=0)
        self.main_area.pack(side="right", fill="both", expand=True)

        self.canvas_container = ctk.CTkScrollableFrame(self.main_area, fg_color="transparent", scrollbar_button_color="#2C6E63")
        self.canvas_container.pack(fill="both", expand=True, padx=50, pady=40)

        self.footer = ctk.CTkFrame(self.main_area, fg_color="#FFFFFF", height=70, corner_radius=0, border_width=1, border_color="#E9EEEC")
        self.footer.pack(side="bottom", fill="x")
        self.footer.pack_propagate(False)

        self.btn_back = ctk.CTkButton(self.footer, text="← Voltar", font=("Arial", 13, "bold"), fg_color="#E9EEEC", text_color="#4B615C", hover_color="#DCE3E0", command=self.prev_step, width=110, height=38)
        self.btn_back.pack(side="left", padx=40)

        self.btn_next = ctk.CTkButton(self.footer, text="Avançar →", font=("Arial", 13, "bold"), fg_color="#2C6E63", text_color="#FFFFFF", hover_color="#1F4F47", command=self.next_step, width=130, height=38)
        self.btn_next.pack(side="right", padx=40)

        self.render_current_view()

    def update_sidebar_styles(self):
        for i, btn in enumerate(self.step_buttons):
            if i == self.current_step:
                btn.configure(fg_color="#20493F", text_color="#F4F8F6", font=("Arial", 13, "bold"))
            else:
                btn.configure(fg_color="transparent", text_color="#9EBBB2", font=("Arial", 13))

        if self.current_step == len(self.steps):
            self.review_btn.configure(fg_color="#20493F", text_color="#F4F8F6", font=("Arial", 13, "bold"))
        else:
            self.review_btn.configure(fg_color="transparent", text_color="#9EBBB2", font=("Arial", 13))

    def jump_to_step(self, idx):
        self.save_current_data()
        self.current_step = idx
        self.render_current_view()

    def render_current_view(self):
        self.update_sidebar_styles()

        for widget in self.canvas_container.winfo_children():
            widget.destroy()

        self.widgets_map = {}
        self.conditional_frames = {}

        if self.current_step < len(self.steps):
            step_title, step_sub, fields = self.steps[self.current_step]

            eyebrow = ctk.CTkLabel(self.canvas_container, text=f"SEÇÃO {self.current_step + 1} DE {len(self.steps)}", font=("Courier", 11, "bold"), text_color="#2C6E63")
            eyebrow.pack(anchor="w", pady=(0, 4))

            title = ctk.CTkLabel(self.canvas_container, text=step_title, font=("Georgia", 24, "bold"), text_color="#16302B")
            title.pack(anchor="w", pady=(0, 4))

            sub = ctk.CTkLabel(self.canvas_container, text=step_sub, font=("Arial", 13), text_color="#4B615C")
            sub.pack(anchor="w", pady=(0, 24))

            i = 0
            while i < len(fields):
                item = fields[i]
                field_id = item[0]
                
                block_frame = ctk.CTkFrame(self.canvas_container, fg_color="transparent")
                block_frame.pack(fill="x", pady=(0, 6))

                self.render_single_field(block_frame, item)

                j = i + 1
                while j < len(fields):
                    sub_item = fields[j]
                    is_conditional = len(sub_item) > 6
                    if is_conditional and sub_item[6] == field_id:
                        sub_field_id = sub_item[0]
                        sub_frame = ctk.CTkFrame(block_frame, fg_color="transparent")
                        self.conditional_frames[sub_field_id] = (sub_frame, field_id, sub_item[7])
                        
                        self.render_single_field(sub_frame, sub_item)

                        parent_val = self.form_data.get(field_id, "")
                        trigger_val = sub_item[7]
                        is_active = False
                        if isinstance(trigger_val, list):
                            is_active = parent_val in trigger_val or (isinstance(parent_val, list) and any(v in trigger_val for v in parent_val))
                        else:
                            is_active = parent_val == trigger_val or (isinstance(parent_val, list) and trigger_val in parent_val)

                        if is_active:
                            sub_frame.pack(fill="x", pady=(4, 0))

                        j += 1
                    else:
                        break
                i = j

            if self.current_step == len(self.steps) - 1:
                lbl_img = ctk.CTkLabel(self.canvas_container, text="Selecione as fotos para embutir no documento Word (cada foto ficará em uma página):", font=("Arial", 13, "bold"), text_color="#16302B")
                lbl_img.pack(anchor="w", pady=(20, 6))

                btn_img = ctk.CTkButton(self.canvas_container, text="📁 Selecionar Fotos...", font=("Arial", 13), fg_color="#E9EEEC", text_color="#173832", hover_color="#DCE3E0", command=self.select_images, width=180, height=36)
                btn_img.pack(anchor="w")

                self.lbl_img_status = ctk.CTkLabel(self.canvas_container, text=f"{len(self.selected_images)} foto(s) selecionada(s)", font=("Arial", 11, "italic"), text_color="#4B615C")
                self.lbl_img_status.pack(anchor="w", pady=(6, 0))

            self.btn_back.configure(state="normal" if self.current_step > 0 else "disabled")
            self.btn_next.configure(text="Ir para revisão →" if self.current_step == len(self.steps) - 1 else "Avançar →", fg_color="#2C6E63")

        else:
            eyebrow = ctk.CTkLabel(self.canvas_container, text="REVISÃO FINAL", font=("Courier", 11, "bold"), text_color="#2C6E63")
            eyebrow.pack(anchor="w", pady=(0, 4))

            title = ctk.CTkLabel(self.canvas_container, text="Revisar e gerar documento", font=("Georgia", 24, "bold"), text_color="#16302B")
            title.pack(anchor="w", pady=(0, 4))

            sub = ctk.CTkLabel(self.canvas_container, text="Confira as informações antes de gerar o documento Word oficial.", font=("Arial", 13), text_color="#4B615C")
            sub.pack(anchor="w", pady=(0, 24))

            for step_title, _, fields in self.steps:
                sec_lbl = ctk.CTkLabel(self.canvas_container, text=step_title.upper(), font=("Arial", 12, "bold"), text_color="#1F4F47")
                sec_lbl.pack(anchor="w", pady=(14, 2))

                for item in fields:
                    field_id, label_text = item[0], item[1]
                    val = self.form_data.get(field_id, "")
                    if isinstance(val, list):
                        val = ", ".join(val)
                    if val:
                        item_lbl = ctk.CTkLabel(self.canvas_container, text=f"• {label_text}: {val}", font=("Arial", 12), text_color="#16302B")
                        item_lbl.pack(anchor="w", padx=10, pady=1)

            btn_gen = ctk.CTkButton(self.canvas_container, text="📄 Baixar Documento Word (.docx) Completo", font=("Arial", 15, "bold"), fg_color="#2C6E63", hover_color="#1F4F47", height=46, command=self.generate_word_document)
            btn_gen.pack(anchor="w", pady=(30, 20))

            self.btn_back.configure(state="normal")
            self.btn_next.configure(text="Gerar Word", fg_color="#2C6E63")

    def render_single_field(self, container, item):
        field_id = item[0]
        label_text = item[1]
        f_type = item[2]
        options = item[3]
        required = item[4]
        highlight = item[5]

        req_str = " *" if required else ""
        lbl_color = "#1F4F47" if highlight else "#16302B"
        lbl = ctk.CTkLabel(container, text=label_text + req_str, font=("Arial", 13, "bold" if (required or highlight) else "normal"), text_color=lbl_color)
        lbl.pack(anchor="w", pady=(4, 2))

        if f_type == "entry":
            ent = ctk.CTkEntry(container, font=("Arial", 13), height=38, fg_color="#FBFCFB", border_color="#2C6E63" if highlight else "#DCE3E0", text_color="#16302B")
            ent.pack(fill="x", pady=(0, 4))
            if field_id in self.form_data:
                ent.insert(0, self.form_data[field_id])
            self.widgets_map[field_id] = ent

        elif f_type == "cpf":
            ent = ctk.CTkEntry(container, font=("Arial", 13), height=38, fg_color="#FBFCFB", border_color="#DCE3E0", text_color="#16302B")
            ent.pack(fill="x", pady=(0, 4))
            if field_id in self.form_data:
                ent.insert(0, self.form_data[field_id])
            ent.bind("<KeyRelease>", lambda e, w=ent: self.formatar_cpf_realtime(w))
            self.widgets_map[field_id] = ent

        elif f_type == "date":
            ent = ctk.CTkEntry(container, font=("Arial", 13), height=38, fg_color="#FBFCFB", border_color="#DCE3E0", text_color="#16302B")
            ent.pack(fill="x", pady=(0, 4))
            if field_id in self.form_data:
                ent.insert(0, self.form_data[field_id])
            ent.bind("<KeyRelease>", lambda e, w=ent: self.formatar_data_realtime(w))
            self.widgets_map[field_id] = ent

        elif f_type == "text":
            txt = ctk.CTkTextbox(container, font=("Arial", 13), height=100, fg_color="#FBFCFB", border_color="#2C6E63" if highlight else "#DCE3E0", text_color="#16302B", border_width=1)
            txt.pack(fill="x", pady=(0, 4))
            if field_id in self.form_data:
                txt.insert("1.0", self.form_data[field_id])
            self.widgets_map[field_id] = txt

        elif f_type == "pills":
            pills_frame = ctk.CTkFrame(container, fg_color="transparent")
            pills_frame.pack(anchor="w", pady=(0, 4))
            
            var = ctk.StringVar(value=self.form_data.get(field_id, ""))
            self.widgets_map[field_id] = var

            for opt in options:
                rb = ctk.CTkRadioButton(
                    pills_frame, text=opt, variable=var, value=opt, font=("Arial", 13), 
                    text_color="#4B615C", fg_color="#2C6E63", hover_color="#1F4F47",
                    command=lambda fid=field_id, v=var: self.on_field_changed(fid, v.get())
                )
                rb.pack(side="left", padx=(0, 20), pady=4)

        elif f_type == "multipills":
            mp_frame = ctk.CTkFrame(container, fg_color="transparent")
            mp_frame.pack(anchor="w", pady=(0, 4))
            
            selected_list = self.form_data.get(field_id, [])
            if not isinstance(selected_list, list):
                selected_list = []
            self.widgets_map[field_id] = selected_list

            for opt in options:
                chk_var = ctk.BooleanVar(value=(opt in selected_list))
                cb = ctk.CTkCheckBox(
                    mp_frame, text=opt, variable=chk_var, font=("Arial", 13),
                    text_color="#4B615C", fg_color="#2C6E63", hover_color="#1F4F47",
                    command=lambda fid=field_id, o=opt, v=chk_var: self.on_multipill_changed(fid, o, v.get())
                )
                cb.pack(anchor="w", pady=3)

    def on_field_changed(self, field_id, value):
        self.form_data[field_id] = value
        self.evaluate_conditionals()

    def on_multipill_changed(self, field_id, option, is_checked):
        if field_id not in self.form_data or not isinstance(self.form_data[field_id], list):
            self.form_data[field_id] = []
        
        if is_checked and option not in self.form_data[field_id]:
            self.form_data[field_id].append(option)
        elif not is_checked and option in self.form_data[field_id]:
            self.form_data[field_id].remove(option)
            
        self.evaluate_conditionals()

    def evaluate_conditionals(self):
        for cond_id, (frame, parent_id, trigger_val) in self.conditional_frames.items():
            parent_val = self.form_data.get(parent_id, "")
            should_show = False
            
            if isinstance(trigger_val, list):
                if isinstance(parent_val, list):
                    should_show = any(v in trigger_val for v in parent_val)
                else:
                    should_show = parent_val in trigger_val
            else:
                if isinstance(parent_val, list):
                    should_show = trigger_val in parent_val
                else:
                    should_show = (parent_val == trigger_val)

            if should_show:
                if not frame.winfo_ismapped():
                    frame.pack(fill="x", pady=(4, 0))
            else:
                if frame.winfo_ismapped():
                    frame.pack_forget()
                if cond_id in self.form_data:
                    self.form_data[cond_id] = "" if not isinstance(self.form_data[cond_id], list) else []

    def formatar_cpf_realtime(self, entry_widget):
        texto = ''.join(filter(str.isdigit, entry_widget.get()))[:11]
        formatado = ""
        if len(texto) > 9:
            formatado = f"{texto[:3]}.{texto[3:6]}.{texto[6:9]}-{texto[9:]}"
        elif len(texto) > 6:
            formatado = f"{texto[:3]}.{texto[3:6]}.{texto[6:]}"
        elif len(texto) > 3:
            formatado = f"{texto[:3]}.{texto[3:]}"
        else:
            formatado = texto
        
        entry_widget.delete(0, "end")
        entry_widget.insert(0, formatado)

    def formatar_data_realtime(self, entry_widget):
        texto = ''.join(filter(str.isdigit, entry_widget.get()))[:8]
        formatado = ""
        if len(texto) > 4:
            formatado = f"{texto[:2]}/{texto[2:4]}/{texto[4:]}"
        elif len(texto) > 2:
            formatado = f"{texto[:2]}/{texto[2:]}"
        else:
            formatado = texto

        entry_widget.delete(0, "end")
        entry_widget.insert(0, formatado)

        if len(texto) == 8:
            try:
                nasc = datetime.strptime(texto, "%d%m%Y")
                hoje = datetime.today()
                idade = hoje.year - nasc.year - ((hoje.month, hoje.day) < (nasc.month, nasc.day))
                if "pac_idade" in self.widgets_map:
                    self.widgets_map["pac_idade"].delete(0, "end")
                    self.widgets_map["pac_idade"].insert(0, str(idade))
                    self.form_data["pac_idade"] = str(idade)
            except ValueError:
                pass

    def save_current_data(self):
        if self.current_step < len(self.steps):
            for field_id, widget in self.widgets_map.items():
                if isinstance(widget, ctk.CTkEntry):
                    self.form_data[field_id] = widget.get().strip()
                elif isinstance(widget, ctk.CTkTextbox):
                    self.form_data[field_id] = widget.get("1.0", "end-1c").strip()
                elif isinstance(widget, ctk.StringVar):
                    self.form_data[field_id] = widget.get().strip()

    def select_images(self):
        files = filedialog.askopenfilenames(title="Selecione as fotos dos exames", filetypes=[("Imagens", "*.jpg *.jpeg *.png *.bmp")])
        if files:
            self.selected_images.extend(files)
            if hasattr(self, 'lbl_img_status'):
                self.lbl_img_status.configure(text=f"{len(self.selected_images)} foto(s) selecionada(s)")

    def next_step(self):
        self.save_current_data()

        if self.current_step < len(self.steps):
            _, _, fields = self.steps[self.current_step]
            for item in fields:
                field_id = item[0]
                label_text = item[1]
                f_type = item[2]
                required = item[4]
                
                if field_id in self.conditional_frames:
                    frame, _, _ = self.conditional_frames[field_id]
                    if not frame.winfo_ismapped():
                        continue

                val = self.form_data.get(field_id, "")
                if isinstance(val, list) and not val:
                    val = ""

                if required and not val:
                    messagebox.showwarning("Campo Obrigatório", f"O campo '{label_text}' é obrigatório.")
                    return
                
                if f_type == "cpf" and val:
                    if not validar_cpf(val):
                        messagebox.showerror("CPF Inválido", f"O CPF informado no campo '{label_text}' não é válido.")
                        return

        if self.current_step <= len(self.steps):
            self.current_step += 1
            if self.current_step > len(self.steps):
                self.generate_word_document()
                self.current_step = len(self.steps)
            else:
                self.render_current_view()

    def prev_step(self):
        self.save_current_data()
        if self.current_step > 0:
            self.current_step -= 1
            self.render_current_view()

    def generate_word_document(self):
        try:
            doc = Document()

            for section in doc.sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)

            # Inserção da Logo no Documento Word (se logo.png existir)
            logo_path = "logo.png"
            if os.path.exists(logo_path):
                try:
                    doc.add_picture(logo_path, width=Inches(1.8))
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

            for step_title, _, fields in self.steps:
                p_sec = doc.add_paragraph()
                r_sec = p_sec.add_run(step_title.upper())
                r_sec.bold = True
                r_sec.font.size = Pt(13)
                r_sec.font.color.rgb = RGBColor(23, 56, 50)
                p_sec.paragraph_format.space_before = Pt(14)
                p_sec.paragraph_format.space_after = Pt(4)

                for item in fields:
                    field_id, label_text = item[0], item[1]
                    
                    if field_id in self.conditional_frames:
                        frame, _, _ = self.conditional_frames[field_id]
                        if not frame.winfo_ismapped():
                            continue

                    val = self.form_data.get(field_id)
                    if isinstance(val, list):
                        val = ", ".join(val)

                    if val:
                        p_field = doc.add_paragraph()
                        p_field.paragraph_format.space_after = Pt(3)
                        r_label = p_field.add_run(f"{label_text}: ")
                        r_label.bold = True
                        r_label.font.size = Pt(11)
                        
                        r_val = p_field.add_run(val)
                        r_val.font.size = Pt(11)

            if self.selected_images:
                for idx, img_path in enumerate(self.selected_images):
                    if os.path.exists(img_path):
                        doc.add_page_break()
                        p_desc = doc.add_paragraph()
                        r_d = p_desc.add_run(f"IMAGEM ANEXADA {idx + 1}: {os.path.basename(img_path)}")
                        r_d.bold = True
                        r_d.font.size = Pt(13)
                        r_d.font.color.rgb = RGBColor(23, 56, 50)
                        p_desc.paragraph_format.space_after = Pt(14)

                        try:
                            doc.add_picture(img_path, width=Inches(5.5))
                        except Exception as img_err:
                            p_err = doc.add_paragraph()
                            p_err.add_run(f"[Erro ao carregar a imagem: {img_err}]").font.color.rgb = RGBColor(181, 80, 46)

            paciente_nome = self.form_data.get("pac_nome", "Paciente").replace(" ", "_")
            file_name = f"Triagem_Vascular_{paciente_nome}.docx"
            
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            if not os.path.exists(desktop_path):
                desktop_path = os.path.expanduser("~")

            self.saved_file_path = os.path.join(desktop_path, file_name)
            doc.save(self.saved_file_path)

            messagebox.showinfo("Sucesso", f"Documento Word gerado com sucesso!\n\nSalvo na sua Área de Trabalho:\n{file_name}")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao gerar o documento:\n{str(e)}")

if __name__ == "__main__":
    app = TriagemVascularApp()
    app.mainloop()
