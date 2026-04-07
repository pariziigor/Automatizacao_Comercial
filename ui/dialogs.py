import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
import requests
import threading
import re

def aplicar_estilo_tabela():
    style = ttk.Style()
    style.theme_use("default")
    
    bg_color = "#2b2b2b"
    text_color = "white"
    selected_color = "#1f538d"
    
    style.configure("Treeview", background=bg_color, foreground=text_color, fieldbackground=bg_color, borderwidth=0, rowheight=25)
    style.configure("Treeview.Heading", background="#1a1a1a", foreground="white", relief="flat")
    style.map("Treeview", background=[("selected", selected_color)])

def janela_verificacao_unificada(parent, todos_placeholders, dados_extraidos):
    win = ctk.CTkToplevel(parent)
    win.title("Passo 1: Conferência Geral")
    win.geometry("950x750")
    win.transient(parent)
    win.grab_set()
    win.lift()
    win.focus_force()
    
    status = {"confirmado": False}
    resultado_final = {}
    widgets_texto = {}
    widgets_servicos = {}
    
    frete_atual = dados_extraidos.get("TIPO_FRETE", "CIF - Por conta do destinatário")
    var_frete = ctk.StringVar(value=frete_atual)

    ctk.CTkLabel(win, text="Revise os dados abaixo", font=("Roboto", 20, "bold")).pack(pady=(20, 5), padx=20, anchor="w")
    ctk.CTkLabel(win, text="Digite o CNPJ e clique na lupa para preencher automático.", text_color="gray").pack(pady=(0, 10), padx=20, anchor="w")

    scroll_frame = ctk.CTkScrollableFrame(win)
    scroll_frame.pack(fill="both", expand=True, padx=20, pady=10)

    def add_section_title(texto):
        ctk.CTkLabel(scroll_frame, text=texto, font=("Roboto", 16, "bold"), text_color=("#3B8ED0", "#FF9800")).pack(anchor="w", pady=(20, 10))
        ctk.CTkFrame(scroll_frame, height=2, fg_color="gray30").pack(fill="x", pady=(0, 10))

    add_section_title("1. OPÇÕES DA PROPOSTA")
    f_opcoes = ctk.CTkFrame(scroll_frame, fg_color="transparent")
    f_opcoes.pack(fill="x", pady=5)
    
    ctk.CTkLabel(f_opcoes, text="Tipo de Frete:").pack(side="left", padx=(0, 10))
    opcoes_frete = ["CIF - Por conta do destinatário", "FOB - Por conta do Cliente"]
    ctk.CTkComboBox(f_opcoes, variable=var_frete, values=opcoes_frete, width=300).pack(side="left")

    lista_servicos = sorted([p for p in todos_placeholders if p.startswith("X_")])
    if lista_servicos:
        add_section_title("2. SELEÇÃO DE SERVIÇOS")
        f_servicos = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        f_servicos.pack(fill="x")
        f_servicos.grid_columnconfigure(0, weight=1)
        f_servicos.grid_columnconfigure(1, weight=1)
        f_servicos.grid_columnconfigure(2, weight=1)

        col, row = 0, 0
        for servico in lista_servicos:
            nome_bonito = servico.replace("X_", "").replace("_", " ").title()
            chk = ctk.CTkCheckBox(f_servicos, text=nome_bonito)
            chk.grid(row=row, column=col, sticky="w", padx=10, pady=10)
            
            if dados_extraidos.get(servico) == "X":
                chk.select()
                
            widgets_servicos[servico] = chk
            col += 1
            if col > 2: col, row = 0, row + 1

    campos_ignorados = ["TIPO_FRETE", "VALOR_TOTAL_PROPOSTA", "ITENS_ORCAMENTO", "ITENS_ESTRUTURAL", "item", "DATA_HOJE"]
    todos_campos_texto = [p for p in todos_placeholders if not p.startswith("X_") and p not in campos_ignorados and not p.startswith("item.")]
    
    grupo_solicitante = [p for p in todos_campos_texto if p.endswith("_SOLICITANTE")]
    grupo_contratante = [p for p in todos_campos_texto if p.endswith("_CONTRATANTE")]
    grupo_outros = [p for p in todos_campos_texto if p not in grupo_solicitante and p not in grupo_contratante]

    def buscar_cnpj_api(entry_cnpj, contexto):
        cnpj_limpo = "".join(filter(str.isdigit, entry_cnpj.get()))
        if len(cnpj_limpo) != 14:
            messagebox.showwarning("Aviso", "O CNPJ deve ter exatamente 14 números.")
            return

        def task():
            try:
                url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj_limpo}"
                headers = {'User-Agent': 'Mozilla/5.0'}
                response = requests.get(url, headers=headers, timeout=10)
                if response.status_code == 200:
                    dados = response.json()
                    if dados.get('status') == 'ERROR':
                        win.after(0, lambda: messagebox.showerror("Erro", dados.get('message', 'CNPJ não encontrado.')))
                        return

                    mapeamento = {
                        "nome": ["NOME_EMPRESA", "RAZAO_SOCIAL"], 
                        "fantasia": ["NOME_FANTASIA"], "cep": ["CEP"],
                        "municipio": ["CIDADE", "MUNICIPIO"], "uf": ["UF", "ESTADO"],
                        "email": ["EMAIL"], "telefone": ["TELEFONE", "CELULAR", "WHATSAPP"], "bairro": ["BAIRRO"]
                    }
                    
                    logradouro = dados.get('logradouro', '')
                    numero = dados.get('numero', '')
                    comp = dados.get('complemento', '')
                    endereco_full = f"{logradouro}, {numero}"
                    if comp: endereco_full += f" - {comp}"

                    def atualizar_gui():
                        count = 0
                        for campo_word, entry_widget in widgets_texto.items():
                            campo_upper = campo_word.upper()
                            if contexto == "FATURAMENTO" and "CONTRATANTE" not in campo_upper: continue
                            if "ENDERECO" in campo_upper and "OBRA" not in campo_upper and "ENTREGA" not in campo_upper:
                                entry_widget.delete(0, "end")
                                entry_widget.insert(0, endereco_full)
                                count += 1
                                continue
                            for chave_api, lista_possiveis in mapeamento.items():
                                for possivel in lista_possiveis:
                                    if possivel in campo_upper:
                                        if "CLIENTE" in campo_upper and chave_api == "nome": continue
                                        valor = dados.get(chave_api)
                                        if valor:
                                            entry_widget.delete(0, "end")
                                            entry_widget.insert(0, str(valor))
                                            count += 1
                        messagebox.showinfo("Sucesso", f"Empresa atualizada! ({count} campos)")
                    win.after(0, atualizar_gui)
                else: win.after(0, lambda: messagebox.showerror("Erro", f"Erro API: {response.status_code}"))
            except Exception as e: win.after(0, lambda: messagebox.showerror("Erro", str(e)))
        threading.Thread(target=task, daemon=True).start()

    def trava_numeros(event, widget):
        if event.keysym in ['Left', 'Right', 'BackSpace', 'Delete']: return
        texto = widget.get()
        limpo = "".join([c for c in texto if c.isdigit() or c in ".-/() "])
        if texto != limpo:
            widget.delete(0, "end")
            widget.insert(0, limpo)

    def trava_letras(event, widget, limite=None):
        if event.keysym in ['Left', 'Right', 'BackSpace', 'Delete']: return
        texto = widget.get()
        limpo = "".join([c for c in texto if c.isalpha() or c.isspace()])
        if limite: limpo = limpo[:limite]
        if texto != limpo:
            widget.delete(0, "end")
            widget.insert(0, limpo)

    def desenhar_campos(lista, titulo):
        if not lista: return
        add_section_title(titulo)
        f_campos = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        f_campos.pack(fill="x", pady=5)
        f_campos.grid_columnconfigure(1, weight=1) 
        f_campos.grid_columnconfigure(3, weight=1) 

        mapa_nomes = {
            "NOME_EMPRESA": "Nome da Empresa / Razão Social", "NOME_CLIENTE": "Nome do Cliente",
            "CNPJ_CPF": "CNPJ ou CPF", "IE": "Inscrição Estadual (I.E.)", "CEP": "CEP",
            "ENDERECO": "Endereço Completo", "CELULAR": "Celular / WhatsApp", "EMAIL": "E-mail",
            "CREA": "Registro CREA / CAU", "ART": "A.R.T.", "RESPONSAVEL_TECNICO": "Responsável Técnico",
            "NUMERO_PROJETO": "Número do Projeto", "TIPO_OBRA": "Tipo de Obra",
            "DESCRICAO_DA_OBRA": "Descrição Detalhada", "ENDERECO_DA_OBRA": "Local da Obra",
            "ARQUIVOS_RECEBIDOS": "Arquivos Recebidos", "DATA_HOJE": "Data de Emissão"
        }

        for i, campo in enumerate(lista):
            valor_auto = dados_extraidos.get(campo, "")
            chave_limpa = campo.replace("_SOLICITANTE", "").replace("_CONTRATANTE", "")
            texto_label = mapa_nomes.get(chave_limpa, chave_limpa.replace("_", " ").title())
            row = i // 2
            col_start = (i % 2) * 2 
            
            ctk.CTkLabel(f_campos, text=texto_label).grid(row=row, column=col_start, sticky="w", padx=5, pady=5)
            
            if "CNPJ" in campo.upper():
                f_cnpj = ctk.CTkFrame(f_campos, fg_color="transparent")
                f_cnpj.grid(row=row, column=col_start+1, sticky="ew", padx=(0, 20), pady=5)
                ent = ctk.CTkEntry(f_cnpj)
                if valor_auto: ent.insert(0, valor_auto)
                ent.pack(side="left", fill="x", expand=True)
                ent.bind("<KeyRelease>", lambda e, w=ent: trava_numeros(e, w))
                tipo_busca = "FATURAMENTO" if "CONTRATANTE" in campo.upper() else "SOLICITANTE"
                ctk.CTkButton(f_cnpj, text="🔍", width=40, command=lambda e=ent, t=tipo_busca: buscar_cnpj_api(e, t)).pack(side="right", padx=(5, 0))
                widgets_texto[campo] = ent
            else:
                ent = ctk.CTkEntry(f_campos)
                if valor_auto: ent.insert(0, valor_auto)
                ent.grid(row=row, column=col_start+1, sticky="ew", padx=(0, 20), pady=5)
                widgets_texto[campo] = ent
                c_upper = campo.upper()
                if any(x in c_upper for x in ["CEP", "CELULAR", "TELEFONE", "WHATSAPP", "CPF"]):
                    ent.bind("<KeyRelease>", lambda e, w=ent: trava_numeros(e, w))
                elif "UF" in c_upper: ent.bind("<KeyRelease>", lambda e, w=ent: trava_letras(e, w, limite=2))
                elif any(x in c_upper for x in ["CIDADE", "MUNICIPIO", "ESTADO"]):
                    ent.bind("<KeyRelease>", lambda e, w=ent: trava_letras(e, w))

    desenhar_campos(grupo_outros, "3. DADOS GERAIS DO PROJETO")
    desenhar_campos(grupo_solicitante, "4. DADOS DO SOLICITANTE")
    desenhar_campos(grupo_contratante, "5. DADOS DE FATURAMENTO")

    def confirmar():
        status["confirmado"] = True
        resultado_final["TIPO_FRETE"] = var_frete.get()
        for k, chk in widgets_servicos.items(): resultado_final[k] = "X" if chk.get() == 1 else "" 
        for k, entry in widgets_texto.items(): resultado_final[k] = entry.get()
        win.destroy()

    ctk.CTkButton(win, text="AVANÇAR", command=confirmar, height=50, fg_color="green", hover_color="darkgreen").pack(fill="x", padx=20, pady=20)
    
    parent.wait_window(win)
    if not status["confirmado"]: return None
    return resultado_final


def janela_projeto_estrutural(parent, dados_anteriores):
    win = ctk.CTkToplevel(parent)
    win.title("Passo 2: Escopo Estrutural")
    win.geometry("900x700")
    win.transient(parent)
    win.grab_set()
    win.lift()
    win.focus_force()

    status = {"acao": None}
    
    lista_itens_estrutural = dados_anteriores.get("ITENS_ESTRUTURAL", []).copy()
    aplicar_estilo_tabela()

    ctk.CTkLabel(win, text="Definição do Escopo Estrutural", font=("Roboto", 20, "bold")).pack(pady=(20, 5))
    ctk.CTkLabel(win, text="Insira os itens de texto livre.", text_color="gray").pack()

    f_input = ctk.CTkFrame(win)
    f_input.pack(fill="x", padx=20, pady=10)

    ctk.CTkLabel(f_input, text="Descrição do Item:").pack(anchor="w", padx=10, pady=(10,0))
    txt_desc = ctk.CTkTextbox(f_input, height=80)
    txt_desc.pack(fill="x", padx=10, pady=5)
    ctk.CTkLabel(f_input, text="Use Enter para criar sub-itens.", font=("Arial", 10), text_color="gray").pack(anchor="w", padx=10)

    f_lista = ctk.CTkFrame(win, fg_color="transparent")
    f_lista.pack(fill="both", expand=True, padx=20, pady=5)

    tree = ttk.Treeview(f_lista, columns=("desc",), show="headings", selectmode="browse")
    tree.heading("desc", text="Itens Adicionados")
    tree.column("desc", anchor="w")
    
    scrollbar = ctk.CTkScrollbar(f_lista, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    for item in lista_itens_estrutural:
        primeira_linha = item.split("\n")[0]
        if len(item.split("\n")) > 1: primeira_linha += " (...)"
        tree.insert("", "end", values=(primeira_linha,), tags=(item,))

    def adicionar_item():
        desc = txt_desc.get("1.0", "end-1c").strip()
        if not desc: return
        primeira_linha = desc.split("\n")[0]
        if len(desc.split("\n")) > 1: primeira_linha += " (...)"
        tree.insert("", "end", values=(primeira_linha,), tags=(desc,))
        lista_itens_estrutural.append(desc)
        txt_desc.delete("1.0", "end")
        txt_desc.focus()

    def remover_item():
        selected = tree.selection()
        if not selected: return
        for item_id in selected:
            idx = tree.index(item_id)
            if idx < len(lista_itens_estrutural): lista_itens_estrutural.pop(idx)
            tree.delete(item_id)

    def voltar():
        status["acao"] = "VOLTAR"
        win.destroy()

    def finalizar():
        status["acao"] = "AVANCAR"
        dados_anteriores["ITENS_ESTRUTURAL"] = lista_itens_estrutural
        win.destroy()

    ctk.CTkButton(f_input, text="Adicionar Item", command=adicionar_item).pack(anchor="e", padx=10, pady=10)
    
    f_footer = ctk.CTkFrame(win, height=50, fg_color="transparent")
    f_footer.pack(fill="x", padx=20, pady=20)
    
    ctk.CTkButton(f_footer, text="VOLTAR", command=voltar, fg_color="gray", hover_color="darkgray", width=100).pack(side="left", padx=(0, 10))
    ctk.CTkButton(f_footer, text="Remover Selecionado", command=remover_item, fg_color="firebrick", hover_color="darkred").pack(side="left")
    ctk.CTkButton(f_footer, text="AVANÇAR", command=finalizar, fg_color="green", hover_color="darkgreen").pack(side="right")

    parent.wait_window(win)
    
    if status["acao"] == "VOLTAR": return "VOLTAR"
    if status["acao"] == "AVANCAR": return dados_anteriores
    return None 


def janela_itens_orcamento(parent, dados_anteriores):
    win = ctk.CTkToplevel(parent)
    win.title("Passo 3: Orçamento")
    win.geometry("900x700")
    win.transient(parent)
    win.grab_set()
    win.lift()
    win.focus_force()

    status = {"acao": None}
    aplicar_estilo_tabela()

    lista_itens = dados_anteriores.get("ITENS_ORCAMENTO", []).copy()
    total_geral_float = 0.0

    def converter_brl_para_float(texto):
        limpo = texto.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
        try: return float(limpo)
        except: return 0.0

    def formatar_moeda_exibicao(valor_float):
        return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    ctk.CTkLabel(win, text="Condições de Pagamento", font=("Roboto", 20, "bold")).pack(pady=(20, 10))
    
    f_input = ctk.CTkFrame(win)
    f_input.pack(fill="x", padx=20, pady=10)
    f_input.grid_columnconfigure(0, weight=1)

    ctk.CTkLabel(f_input, text="Descrição:").grid(row=0, column=0, sticky="w", padx=10, pady=(10,0))
    ctk.CTkLabel(f_input, text="Valor (R$):").grid(row=0, column=1, sticky="w", padx=10, pady=(10,0))
    
    entry_desc = ctk.CTkEntry(f_input)
    entry_desc.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
    
    entry_valor = ctk.CTkEntry(f_input, width=150)
    entry_valor.grid(row=1, column=1, sticky="ew", padx=10, pady=(0, 10))
    entry_valor.insert(0, "0,00")

    def formatar_moeda_ao_digitar(event):
        valor_atual = entry_valor.get()
        apenas_numeros = "".join(filter(str.isdigit, valor_atual))
        if not apenas_numeros: apenas_numeros = "0"
        valor_float = int(apenas_numeros) / 100
        formatado = f"{valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        entry_valor.delete(0, "end")
        entry_valor.insert(0, formatado)

    entry_valor.bind("<KeyRelease>", formatar_moeda_ao_digitar)

    f_lista = ctk.CTkFrame(win, fg_color="transparent")
    f_lista.pack(fill="both", expand=True, padx=20, pady=5)

    tree = ttk.Treeview(f_lista, columns=("desc", "valor"), show="headings")
    tree.heading("desc", text="Descrição")
    tree.heading("valor", text="Valor")
    tree.column("desc", width=500)
    tree.column("valor", width=150, anchor="e")

    scrollbar = ctk.CTkScrollbar(f_lista, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    for item in lista_itens:
        val_float = converter_brl_para_float(item["valor"])
        total_geral_float += val_float
        tree.insert("", "end", values=(item["descricao"], item["valor"]))

    lbl_total = ctk.CTkLabel(win, text=f"TOTAL GERAL: {formatar_moeda_exibicao(total_geral_float)}", font=("Roboto", 24, "bold"), text_color="green")
    lbl_total.pack(pady=10)

    def adicionar_item():
        nonlocal total_geral_float
        desc = entry_desc.get().strip()
        valor_txt = entry_valor.get().strip()
        if not desc: return
        try:
            valor_float = converter_brl_para_float(valor_txt)
            if valor_float <= 0:
                messagebox.showwarning("Atenção", "O valor deve ser maior que zero.")
                return
            tree.insert("", "end", values=(desc, formatar_moeda_exibicao(valor_float)))
            lista_itens.append({"descricao": desc, "valor": formatar_moeda_exibicao(valor_float)})
            total_geral_float += valor_float
            lbl_total.configure(text=f"TOTAL GERAL: {formatar_moeda_exibicao(total_geral_float)}")
            entry_desc.delete(0, "end")
            entry_valor.delete(0, "end")
            entry_valor.insert(0, "0,00")
            entry_desc.focus()
        except Exception as e: messagebox.showerror("Erro", f"Valor inválido: {e}")

    def remover_item():
        nonlocal total_geral_float
        selected = tree.selection()
        if not selected: return
        for item_id in selected: tree.delete(item_id)
        lista_itens.clear()
        total_geral_float = 0.0
        for child in tree.get_children():
            vals = tree.item(child)['values']
            val_float = converter_brl_para_float(str(vals[1]))
            lista_itens.append({"descricao": vals[0], "valor": vals[1]})
            total_geral_float += val_float
        lbl_total.configure(text=f"TOTAL GERAL: {formatar_moeda_exibicao(total_geral_float)}")

    def voltar():
        status["acao"] = "VOLTAR"
        win.destroy()

    def finalizar():
        status["acao"] = "AVANCAR"
        dados_anteriores["ITENS_ORCAMENTO"] = lista_itens
        dados_anteriores["VALOR_TOTAL_PROPOSTA"] = formatar_moeda_exibicao(total_geral_float)
        win.destroy()

    ctk.CTkButton(f_input, text="Adicionar (+)", command=adicionar_item, width=120).grid(row=1, column=2, padx=10, pady=(0, 10))
    
    f_footer = ctk.CTkFrame(win, fg_color="transparent")
    f_footer.pack(fill="x", padx=20, pady=20)
    
    ctk.CTkButton(f_footer, text="VOLTAR", command=voltar, fg_color="gray", hover_color="darkgray", width=100).pack(side="left", padx=(0, 10))
    ctk.CTkButton(f_footer, text="Remover Item", command=remover_item, fg_color="firebrick", hover_color="darkred").pack(side="left")
    ctk.CTkButton(f_footer, text="GERAR PROPOSTA", command=finalizar, fg_color="green", hover_color="darkgreen", height=40).pack(side="right", fill="x", expand=True, padx=(20, 0))

    parent.wait_window(win)
    
    if status["acao"] == "VOLTAR": return "VOLTAR"
    if status["acao"] == "AVANCAR": return dados_anteriores
    return None 