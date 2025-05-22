# Automatização_docx
Nesse projeto realizei-a  automação de um processo extremamente respetivo. Ideia surgiu após uma funcioaria deixar o cargo que executava essa função
após tomei as atividades e desenvolvi esse projeto de que manipula arquivos em **docx e xmls**. <br>
Somente alguns pontos dos arquivos são alterados pois tratam-se de documentos padrão, logo somente as informações de cada cliente devem ser alteradas.

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
        "nome_cliente": "FRANCISCO LUCAS ALMEIDA DOS ANJOS",
        "cpf": "074.470.553-36",
        "rg": "20075075649",
        "endereco_completo": "Setor Quixoa, 0 - Barro alto, Iguatu, CE - CEP: 63500-005",
        "MODULOS": "LEAPTON SOLAR LP182-M-66-NH-570W",
        "QUANT_MODULOS": "13",
        "INVERSOR": "SOLPLANET ASW7300-S",
        "QUANT_INVERSOR": "1",
        "valor_projeto": "R$ 20.000,00",
        "potencia_kwp": "7,41 kWp",
        "forma_pagamento": "Pagamento 100% Bradesco",
        "data_assinatura": datetime.date.today().strftime("%d/%m/%Y"),
        ##Dicionario procuração
        "est_civil": "Casado",
        "profissao":"Empresario",
        "email":"lucas96almeida@icloud.com",
        "contato":"88 98130-3357",
        "end.rua":"Setor Quixoa, 0",
        "end.numero":"0",
        "end.bairro":"Barro alto",
        "end.cidade":"Iguatu, CE",
        "end.cep":"CEP: 63500-005",
        ##Dicionario cadUni
        "vendedor":"Ildernando",
        "cnpj":" ",
        "data_nascimento":"14/05/1996"
    }

    
