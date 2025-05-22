# Automatização_docx
Nesse projeto realizei-a  automação de um processo extremamente respetivo. Ideia surgiu após uma funcioaria deixar o cargo que executava essa função
após tomei as atividades e desenvolvi esse projeto de que manipula arquivos em **docx e xmls**. <br>
Somente alguns pontos dos arquivos são alterados pois tratam-se de documentos padrão, logo somente as informações de cada cliente devem ser alteradas.
Os docx também serão salvos em PDF.

Projeto desenvolvido no Jupyter Notebook

**Formatando Fontes**

    def formatar_run(run):
      run.font.name = 'Arial'
      run.font.size = Pt(12)
      r = run._element
      r.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

**Importando bibliotecas** 

    from docx import Document 
    from docx.shared import Pt 
    from docx.shared import Inches 
    from docx.oxml.ns import qn 
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT 
    import datetime 
    import re 
    import unicodedata 
    from docx2pdf import convert 
    import pandas as pd 
    from openpyxl import load_workbook 

**Percorrendo todos os parafgrafos dos documentos**

    def substituir_campos(doc, campos):
        for paragrafo in doc.paragraphs:
            for chave, valor in campos.items():
                marcador = f"{{{{{chave}}}}}"
                if marcador in paragrafo.text:
                    novo_texto = paragrafo.text = paragrafo.text.replace(marcador, str(valor))
                    paragrafo.clear()
                    run = paragrafo.add_run(novo_texto)
                    formatar_run(run)
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for chave, valor in campos.items():
                        marcador = f"{{{{{chave}}}}}"
                        if marcador in celula.text:
                            celula.text = celula.text.replace(marcador, str(valor))
                            if marcador in paragrafo.text:
                                novo_texto = paragrafo.text.replace(marcador, str(valor))
                                paragrafo.clear()
                                run = paragrafo.add_run(novo_texto)
                                formatar_run(run)
    
        return doc
    
    def limpar_nome(nome):
        # Remove acentos e espaços, substitui por underline
        nome_sem_acentos = unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('utf-8')
        nome_formatado = re.sub(r'\W+', ' ', nome_sem_acentos)
        return nome_formatado.upper()
        def substituir_em_planilha(modelo_excel, dados, nome_cliente):
    wb = load_workbook(modelo_excel)
    
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for chave, valor in dados.items():
                        marcador = f"{{{{{chave}}}}}"
                        if marcador in cell.value:
                            cell.value = cell.value.replace(marcador, str(valor))
    
    nome_arquivo_excel = f" Cadastro Unificado - {limpar_nome(nome_cliente)}.xlsx"
    wb.save(nome_arquivo_excel)
    print(f"Planilha preenchida salva como: {nome_arquivo_excel}")

  **Dicionarios com as informações do clientes**

    dados_cliente = {
        ##dicinario contrato comercial
        "nome_cliente": "XXXXXXXXX",
        "cpf": "XXXXXXXXX",
        "rg": "XXXXXXXXX",
        "endereco_completo": "XXXXXXXXX",
        "MODULOS": "XXXXXXXXX",
        "QUANT_MODULOS": "XXXXXXXXX",
        "INVERSOR": "XXXXXXXXX",
        "QUANT_INVERSOR": "XXXXXXXXX",
        "valor_projeto": "R$ XXXXXXXXX",
        "potencia_kwp": "XXXXXXXXX",
        "forma_pagamento": "XXXXXXXXX",
        "data_assinatura": datetime.date.today().strftime("%d/%m/%Y"),
        ##Dicionario procuração
        "est_civil": "XXXXXXXXX",
        "profissao":"XXXXXXXXX",
        "email":"XXXXXXXXX",
        "contato":"XXXXXXXXX",
        "end.rua":"XXXXXXXXX",
        "end.numero":"XXXXXXXXX",
        "end.bairro":"XXXXXXXXX",
        "end.cidade":"XXXXXXXXX",
        "end.cep":"CEP: XXXXXXXXX",
        ##Dicionario cadUni
        "vendedor":"XXXXXXXXX",
        "cnpj":" XXXXXXXXX",
        "data_nascimento":"XXXXXXXXX"
    }

**Salvando Docx Contrato**

    nome_arquivo = f"CONTRATO COMERCIAL - {limpar_nome(dados_cliente['nome_cliente'])} + MINUTA TERMO DE CIENCIA E IMAGEM.docx"
    doc_modelo = Document("Contrato_Comercial.docx")
    doc_preenchido = substituir_campos(doc_modelo, dados_cliente)
    doc_preenchido.save(nome_arquivo)
    print(f"Contrato salvo como: {nome_arquivo}")

**Salvando Docx Procuração**

    doc_modelo_2 = Document("Modelo_Proc_PF.docx")

    # Substitui os campos
    doc_preenchido_2 = substituir_campos(doc_modelo_2, dados_cliente)

    # Nome do arquivo baseado no nome do novo cliente
    nome_arquivo_2 = f"Procuração Pessoa Fisica - {limpar_nome(dados_cliente['nome_cliente'])}.docx"
    doc_preenchido_2.save(nome_arquivo_2)

**Gerando PDF Contrato e Procuração**

    ##PDF Contrato
    convert(nome_arquivo)
    nome_pdf = nome_arquivo.replace(".docx", ".pdf")
    print(f"PDF também salvo como: {nome_pdf}")

    #PDF Procuração
    convert(nome_arquivo_2)
    print(f"Contrato salvo como: {nome_arquivo_2}")
    print(f"PDF salvo como: {nome_arquivo_2.replace('.docx', '.pdf')}")

**Salvando Planilha de cadastro unificado**
    
    substituir_em_planilha("Cadastro_Unificado.xlsx", dados_cliente, dados_cliente["nome_cliente"])    
