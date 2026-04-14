import customtkinter as ctk
import tkinter as tk
import json
import os
import threading
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from services.utils import resource_path
from services import reader, parser, generator
from ui.dialogs import janela_verificacao_unificada, janela_itens_orcamento, janela_projeto_estrutural

class GeradorPropostasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Propostas Comerciais - AULEVI")
        self.root.geometry("700x600")

        try:
            self.root.iconbitmap(resource_path("icone.ico"))
        except:
            pass
        
        config = self._carregar_config()

        self.path_modelo = resource_path("proposta_modelo.docx")
        print(f"Usando modelo interno: {self.path_modelo}") 

        self.path_solicitante = tk.StringVar()
        self.path_faturamento = tk.StringVar()
        self.path_saida = tk.StringVar(value=config.get("saida", ""))
        
        self.dados_finais = {}
        
        self._setup_ui()

    def _setup_ui(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(2, weight=1)

        lbl_titulo = ctk.CTkLabel(self.root, text="Gerador de Propostas - AULEVI", font=("Roboto", 24, "bold"))
        lbl_titulo.pack(pady=(20, 10))

        f_arquivos = ctk.CTkFrame(self.root)
        f_arquivos.pack(fill="x", padx=20, pady=10)
        f_arquivos.grid_columnconfigure(1, weight=1) 

        ctk.CTkLabel(f_arquivos, text="Arquivos de Entrada", font=("Roboto", 14, "bold")).grid(row=0, column=0, sticky="w", padx=15, pady=10)
        
        ctk.CTkLabel(f_arquivos, text="Modelo Word:").grid(row=1, column=0, sticky="w", padx=15, pady=5)
        ctk.CTkLabel(f_arquivos, text="Padrão Interno (Automático)", text_color="orange").grid(row=1, column=1, sticky="w", padx=5, pady=5)

        div = ctk.CTkFrame(f_arquivos, height=2, fg_color="gray30")
        div.grid(row=4, column=0, columnspan=4, sticky="ew", pady=15, padx=10)
        
        ctk.CTkLabel(f_arquivos, text="Salvar em:").grid(row=5, column=0, sticky="w", padx=15)
        entry_saida = ctk.CTkEntry(f_arquivos, textvariable=self.path_saida)
        entry_saida.grid(row=5, column=1, sticky="ew", padx=5, pady=5)
        ctk.CTkButton(f_arquivos, text="Selecionar", width=100, command=self._buscar_pasta).grid(row=5, column=2, padx=15)

        self.btn_processar = ctk.CTkButton(
            self.root, text="INICIAR PROCESSO", command=self.fluxo_principal, 
            state="normal", height=50, font=("Roboto", 16, "bold"),
            fg_color="green", hover_color="darkgreen"
        )
        self.btn_processar.pack(fill="x", padx=20, pady=20)
        
        self.log_text = ctk.CTkTextbox(self.root)
        self.log_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.log_text.configure(state="disabled")

        self.lbl_progresso = ctk.CTkLabel(self.root, text="", text_color="orange")
        self.lbl_progresso.pack_forget()
        
        self.bar_progresso = ctk.CTkProgressBar(self.root, mode="determinate", height=15)
        self.bar_progresso.set(0)
        self.bar_progresso.pack_forget()

    def _criar_seletor(self, parent, label, var, row):
        ctk.CTkLabel(parent, text=label).grid(row=row, column=0, sticky="w", padx=15, pady=5)
        ctk.CTkEntry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=5, pady=5)
        ctk.CTkButton(parent, text="Arquivo...", width=100, command=lambda: self._buscar_arquivo(var)).grid(row=row, column=2, padx=15, pady=5)

    def _buscar_arquivo(self, var_target):
        path = filedialog.askopenfilename()
        if path: var_target.set(path)
    
    def _buscar_pasta(self):
        path = filedialog.askdirectory()
        if path:
            self.path_saida.set(path)
            self._salvar_config()

    def log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"> {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _carregar_config(self):
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r") as f: return json.load(f)
            except: return {}
        return {}

    def _salvar_config(self):
        try:
            with open("config.json", "w") as f:
                json.dump({"saida": self.path_saida.get()}, f, indent=4)
        except: pass

    def fluxo_principal(self):
        try:
            self.log("1. Analisando Modelo...")
            placeholders = reader.extrair_placeholders_modelo(self.path_modelo)
            
            dados_auto = {}
            if self.path_solicitante.get():
                self.log("2. Lendo dados do Solicitante...")
                txt_sol = reader.ler_documento_cliente(self.path_solicitante.get())
                dados_auto = parser.processar_dados(placeholders, txt_sol)
            else:
                self.log("2. Ficha Solicitante não informada. Preenchimento manual...")

            if self.path_faturamento.get():
                self.log("Lendo dados de Faturamento...")
                txt_fat = reader.ler_documento_cliente(self.path_faturamento.get())
                dados_fat = parser.processar_dados(placeholders, txt_fat, sufixo_filtro="_CONTRATANTE")
                dados_auto.update(dados_fat)

            passo_atual = 1
            dados_acumulados = dados_auto.copy()

            while passo_atual <= 3:
                if passo_atual == 1:
                    self.log("3. Revisão (Passo 1/3)...")
                    resultado = janela_verificacao_unificada(self.root, placeholders, dados_acumulados)
                    if resultado is None:
                        self.log(">> PROCESSO CANCELADO PELO USUÁRIO <<")
                        self.btn_processar.configure(state="normal")
                        return
                    
                    dados_acumulados.update(resultado)
                    passo_atual = 2

                elif passo_atual == 2:
                    self.log("4. Estrutural (Passo 2/3)...")
                    resultado = janela_projeto_estrutural(self.root, dados_acumulados)
                    if resultado is None:
                        self.log(">> PROCESSO CANCELADO PELO USUÁRIO <<")
                        self.btn_processar.configure(state="normal")
                        return
                    if resultado == "VOLTAR":
                        passo_atual = 1
                        continue
                    
                    dados_acumulados = resultado
                    passo_atual = 3

                elif passo_atual == 3:
                    self.log("5. Orçamento (Passo 3/3)...")
                    resultado = janela_itens_orcamento(self.root, dados_acumulados)
                    if resultado is None:
                        self.log(">> PROCESSO CANCELADO PELO USUÁRIO <<")
                        self.btn_processar.configure(state="normal")
                        return
                    if resultado == "VOLTAR":
                        passo_atual = 2
                        continue
                    
                    dados_acumulados = resultado
                    break 
            
            dados_finais = dados_acumulados

            agora = datetime.now()
            meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
            dados_finais["DATA_HOJE"] = f"{agora.day} de {meses[agora.month - 1]} de {agora.year}"
            
            self.log("6. Iniciando geração...")
            
            self.lbl_progresso.configure(text="0% - Iniciando...")
            self.lbl_progresso.pack(before=self.log_text, pady=(0, 5))
            self.bar_progresso.pack(before=self.log_text, fill="x", padx=50, pady=(0, 20))
            self.bar_progresso.set(0)
            self.btn_processar.configure(state="disabled") 

            def atualizar_gui(pct, mensagem):
                self.bar_progresso.set(pct / 100)
                self.lbl_progresso.configure(text=f"{pct}% - {mensagem}")
                self.root.update_idletasks()

            def tarefa_pesada():
                try:
                    generator.gerar_arquivos(
                        self.path_modelo, 
                        dados_finais, 
                        self.path_saida.get(), 
                        callback_progresso=atualizar_gui
                    )
                    self.root.after(0, finalizar_sucesso)
                except Exception as e:
                    self.root.after(0, lambda: finalizar_erro(str(e)))

            def finalizar_sucesso():
                self.lbl_progresso.configure(text="100% - Concluído!")
                self.bar_progresso.set(1)
                self.btn_processar.configure(state="normal")
                self.log("Concluído com sucesso.")
                messagebox.showinfo("Sucesso", "Arquivos gerados com sucesso!")
                self.root.after(3000, lambda: self.bar_progresso.pack_forget())
                self.root.after(3000, lambda: self.lbl_progresso.pack_forget())

            def finalizar_erro(msg_erro):
                self.btn_processar.configure(state="normal")
                self.log(f"ERRO: {msg_erro}")
                messagebox.showerror("Erro", msg_erro)

            threading.Thread(target=tarefa_pesada, daemon=True).start()

        except Exception as e:
            self.log(f"ERRO GERAL: {e}")
            messagebox.showerror("Erro", str(e))