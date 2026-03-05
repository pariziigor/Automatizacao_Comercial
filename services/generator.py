from docxtpl import DocxTemplate
from docx2pdf import convert
from datetime import datetime
import os
import sys
import re

def gerar_arquivos(path_modelo, dados, pasta_saida, callback_progresso=None):
    
    def reportar(pct, msg):
        if callback_progresso:
            callback_progresso(pct, msg)

    reportar(10, "Carregando modelo Word...")
    try:
        doc = DocxTemplate(path_modelo)
    except Exception as e:
        raise Exception(f"Erro ao abrir o modelo Word: {str(e)}")

    reportar(30, "Processando e padronizando textos...")
    
    # --- LÓGICA DE PADRONIZAÇÃO (MAIÚSCULAS/MINÚSCULAS) ---
    dados_limpos = {}
    for chave, valor in dados.items():
        if valor is None:
            dados_limpos[chave] = ""
        elif isinstance(valor, list):
            # Mantém as listas (Orçamento e Estrutural) intactas
            dados_limpos[chave] = valor
        else:
            texto = str(valor).strip()
            chave_upper = chave.upper()
            
            # Regras de formatação por tipo de campo
            if "EMAIL" in chave_upper:
                dados_limpos[chave] = texto.lower()
                
            elif any(x in chave_upper for x in ["UF", "ESTADO", "CEP", "CNPJ", "CPF", "IE"]):
                dados_limpos[chave] = texto.upper()
                
            elif any(x in chave_upper for x in ["DATA", "VALOR", "FRETE", "X_"]):
                # Mantém originais: Datas, Valores em R$, Frete e check de serviços
                dados_limpos[chave] = texto
                
            else:
                # Aplica o padrão: Primeira Maiúscula, o resto minúscula
                texto_formatado = texto.title()
                
                # Ajuste fino: transforma "De", "Da", "Do" em minúsculas
                preposicoes = [" De ", " Da ", " Do ", " Das ", " Dos ", " E "]
                for prep in preposicoes:
                    texto_formatado = texto_formatado.replace(prep, prep.lower())
                    
                dados_limpos[chave] = texto_formatado
    
    # Renderiza o documento
    try:
        doc.render(dados_limpos)
    except Exception as e:
        raise Exception(f"Erro ao preencher o Word (Tags incorretas?): {str(e)}")

    # --- Lógica de Nome do Arquivo ---
    nome_cliente = dados.get("NOME_EMPRESA_SOLICITANTE") or \
                   dados.get("NOME_EMPRESA") or \
                   dados.get("RAZAO_SOCIAL") or \
                   dados.get("NOME_CLIENTE") or \
                   "Cliente"
                   
    # Padroniza o nome do cliente no nome do arquivo
    nome_cliente = nome_cliente.title() 
    
    num_projeto = dados.get("NUMERO_PROJETO") or "000"
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    
    nome_base = f"Proposta {num_projeto} - {nome_cliente} - {data_hoje}"
    nome_final = re.sub(r'[<>:"/\\|?*]', '', str(nome_base)).strip()

    path_word = os.path.join(pasta_saida, f"{nome_final}.docx")
    path_pdf = os.path.join(pasta_saida, f"{nome_final}.pdf")

    reportar(50, f"Salvando DOCX...")
    doc.save(path_word)

    reportar(70, "Convertendo para PDF...")
    try:
        if sys.platform == "win32":
            convert(path_word, path_pdf)
        else:
            pass 
    except Exception as e:
        print(f"Erro PDF (Ignorado): {e}")
    
    reportar(100, "Concluído!")
    
    return path_word, path_pdf