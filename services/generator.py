from docxtpl import DocxTemplate
from datetime import datetime
import os
import sys
import re
import pythoncom
import win32com.client

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
    
    dados_limpos = {}
    for chave, valor in dados.items():
        if valor is None:
            dados_limpos[chave] = ""
        elif isinstance(valor, list):
            dados_limpos[chave] = valor
        else:
            texto = str(valor).strip()
            chave_upper = chave.upper()
            
            if "EMAIL" in chave_upper:
                dados_limpos[chave] = texto.lower()
                
            elif "CNPJ" in chave_upper:
                # Extrai apenas os números do que o usuário digitou
                apenas_numeros = re.sub(r'\D', '', texto)
                
                if len(apenas_numeros) == 14:
                    # Formata no padrão XX.XXX.XXX/XXXX-XX
                    cnpj_formatado = f"{apenas_numeros[:2]}.{apenas_numeros[2:5]}.{apenas_numeros[5:8]}/{apenas_numeros[8:12]}-{apenas_numeros[12:]}"
                    dados_limpos[chave] = cnpj_formatado
                elif len(apenas_numeros) == 11:
                    # Se por acaso for um CPF (11 dígitos): XXX.XXX.XXX-XX
                    cpf_formatado = f"{apenas_numeros[:3]}.{apenas_numeros[3:6]}.{apenas_numeros[6:9]}-{apenas_numeros[9:]}"
                    dados_limpos[chave] = cpf_formatado
                else:
                    # Se estiver incompleto, apenas deixa em maiúsculo
                    dados_limpos[chave] = texto.upper()
                    
            elif any(x in chave_upper for x in ["UF", "ESTADO", "CEP", "CPF", "IE"]):
                dados_limpos[chave] = texto.upper()
                
            elif any(x in chave_upper for x in ["DATA", "VALOR", "FRETE", "X_"]):
                dados_limpos[chave] = texto
                
            else:
                texto_formatado = texto.title()
                preposicoes = [" De ", " Da ", " Do ", " Das ", " Dos ", " E "]
                for prep in preposicoes:
                    texto_formatado = texto_formatado.replace(prep, prep.lower())
                dados_limpos[chave] = texto_formatado
    
    try:
        doc.render(dados_limpos)
    except Exception as e:
        raise Exception(f"Erro ao preencher o Word: {str(e)}")

    nome_cliente = dados.get("NOME_EMPRESA_SOLICITANTE") or \
                   dados.get("NOME_EMPRESA") or \
                   dados.get("RAZAO_SOCIAL") or \
                   dados.get("NOME_CLIENTE") or \
                   "Cliente"
                   
    nome_cliente = nome_cliente.title() 
    
    num_projeto = dados.get("NUMERO_PROJETO") or "000"
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    
    nome_base = f"Proposta {num_projeto} - {nome_cliente} - {data_hoje}"
    nome_final = re.sub(r'[<>:"/\\|?*]', '', str(nome_base)).strip()

    path_word = os.path.abspath(os.path.join(pasta_saida, f"{nome_final}.docx"))
    path_pdf = os.path.abspath(os.path.join(pasta_saida, f"{nome_final}.pdf"))

    reportar(50, f"Salvando DOCX...")
    doc.save(path_word)

    reportar(70, "Convertendo para PDF...")
    try:
        if sys.platform == "win32":
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            
            try:
                documento = word.Documents.Open(path_word)
                documento.SaveAs(path_pdf, FileFormat=17)
                documento.Close()
            finally:
                word.Quit()
                
            pythoncom.CoUninitialize()
        else:
            pass 
    except Exception as e:
        reportar(70, f"Aviso PDF: {str(e)}")
        print(f"Erro PDF: {e}")
    
    reportar(100, "Concluído!")
    
    return path_word, path_pdf