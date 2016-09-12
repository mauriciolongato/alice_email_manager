# -*- coding: latin-1 -*-
from xlrd import open_workbook
import sqlite3 as sql
import os
import time
import re

# Li linha a linha do excel
def readRows(sheet):
    # using list comprehension
    return [sheet.row_values(idx) for idx in range(sheet.nrows)]

# Dado uma string, retorna somente os numeros
def trata_ncm(ncm_string):
    # print "string_ncm: ", type(ncm_string), ncm_string
    ncm = re.findall(r'\d+', ncm_string)
    # print "imprime ncms tratados: ", ncm
    return ncm

# Checar se existe arquivo com analysis_status NULL
conn = sql.connect('C:/Users/mauricio.longato/PycharmProjects/gmail_api/email_tracker.db')
procs_query = conn.execute('''select * from download_inventory where unzip_status = 1 AND analysis_status is null;''')
procs = procs_query.fetchall()
conn.close()

# Obtem pasta corrente
current_path = os.path.dirname(os.path.realpath(__file__))
attachment_files = current_path + "/unziped_attachments"

# Inicia processo de leitura do arquivo
n_arquivo = 1
data_arq = []
for file in procs:
    # Constroi o endereco do arquivo
    file_name_dir = file[1][:len(file[1])-4]
    file_name = file[2]
    file_address = attachment_files+'/'+file_name_dir+'/'+file_name

    # Abre o arquivo
    book = open_workbook(file_address, on_demand=True)
    # print("arquivo numero: ", n_arquivo, " - ", file)
    data = []
    try:
        # O header dos dados fica na primeira aba do arquivo extraido
        sheet_1 = book.sheet_by_name('Aliceweb_parte_1')
        nrow = 0
        row_inicio = None
        # o Header deve ser construido somente com as infomarções da primeira aba
        for cell in sheet_1.col(0):
            # Inicio dos dados (sem cabeçalho)
            if 'Código' in cell.value:
                row_inicio = nrow
            nrow += 1

        # Separa o header
        header_raw = sheet_1.col(0)[0:row_inicio]

        # Inicia a classificacao do arquivo do aliceweb - valores default
        email_id = file[0]
        attachment_file_name = file_name
        file_address = file_address = attachment_files+'/'+file_name_dir
        date_read = str(time.strftime("%d/%m/%Y %H:%M:%S"))
        sentido_trade = None
        tipo_ncm = None
        modo_ncm = None
        lista_ncm = None
        bloco = None
        pais = None
        uf = None
        porto = None
        via = None
        primeiroDetalhamento = None
        segundoDetalhamento = None
        P1 = None
        P2 = None
        P3 = None
        P4 = None
        P5 = None
        P6 = None

        # Inicia a classificacao das informacoes
        for row in header_raw:
            # Sentido imp ou exp
            if 'IMPORTAÇÃO' in row.value:
                sentido_trade = 'Importação'

            if 'EXPORTAÇÃO' in row.value:
                sentido_trade = 'Exportação'

            # Caso o metodo de input foi ncm por ncm - Cesta de produtos
            if 'Cesta de Produtos' in row.value.split(':  ')[0]:
                lista_ncm = str(trata_ncm(row.value.split(':  ')[1]))
                tipo_ncm = str(len(lista_ncm[0]))
                modo_ncm = "unitario"

            # Caso o metodo de input foi intervalo de NCMs - Capítulo -
            if 'Capítulo -' in row.value.split(':  ')[0]:
                lista_ncm = str(trata_ncm(row.value.split(':  ')[1]))
                tipo_ncm = str(len(lista_ncm[0]))
                modo_ncm = "intervalo"

            # Bloco economico
            if 'Bloco Econômico' in row.value.split(':  ')[0]:
                bloco = row.value.split(':  ')[1]

            # Pais
            if 'País' in row.value.split(':  ')[0]:
                bloco = row.value.split(':  ')[1]

            # UF
            if 'UF' in row.value.split(':  ')[0]:
                uf = row.value.split(':  ')[1]

            # Porto
            if 'Porto' in row.value.split(':  ')[0]:
                porto = row.value.split(':  ')[1]

            # Via
            if 'Via' in row.value.split(':  ')[0]:
                via = row.value.split(':  ')[1]

            # Primeiro Detalhamento
            if "Primeiro detalhamento" in row.value.split(':  ')[0]:
                primeiroDetalhamento = row.value.split(':  ')[1]

            # Segundo Detalhamento
            if "Segundo detalhamento" in row.value.split(':  ')[0]:
                segundoDetalhamento = row.value.split(':  ')[1]

            # P1
            if "P1" in row.value.split(':  ')[0]:
                P1 = trata_ncm(row.value.split(':  ')[1])
                P1 = str(P1[1])+str(P1[0])+" - "+str(P1[3])+str(P1[2])

            if "P2" in row.value.split(':  ')[0]:
                P2 = trata_ncm(row.value.split(':  ')[1])
                P2 = str(P2[1]) + str(P2[0]) + " - " + str(P2[3]) + str(P2[2])

            if "P3" in row.value.split(':  ')[0]:
                P3 = trata_ncm(row.value.split(':  ')[1])
                P3 = str(P3[1]) + str(P3[0]) + " - " + str(P3[3]) + str(P3[2])

            if "P4" in row.value.split(':  ')[0]:
                P4 = trata_ncm(row.value.split(':  ')[1])
                P4 = str(P4[1]) + str(P4[0]) + " - " + str(P4[3]) + str(P4[2])

            if "P5" in row.value.split(':  ')[0]:
                P5 = trata_ncm(row.value.split(':  ')[1])
                P5 = str(P5[1]) + str(P5[0]) + " - " + str(P5[3]) + str(P5[2])

            if "P6" in row.value.split(':  ')[0]:
                P6 = trata_ncm(row.value.split(':  ')[1])
                P6 = str(P6[1]) + str(P6[0]) + " - " + str(P6[3]) + str(P6[2])

        # Coloca informacoes em formato de vetor
        info_row = [(
        email_id, attachment_file_name, file_address, date_read, sentido_trade, str(tipo_ncm), modo_ncm, str(lista_ncm), bloco,
        pais, uf, porto, via, primeiroDetalhamento, segundoDetalhamento, P1, P2, P3, P4, P5, P6)]

        print(type(email_id), email_id)
        print(type(attachment_file_name), attachment_file_name)
        print(type(file_address), file_address)
        print(type(date_read), date_read)
        print(type(sentido_trade), sentido_trade)
        print(type(tipo_ncm), tipo_ncm)
        print(type(modo_ncm), modo_ncm)
        print(type(lista_ncm), lista_ncm)
        print(type(bloco), bloco)
        print(type(pais), pais)
        print(type(uf), uf)
        print(type(porto), porto)
        print(type(via), via)
        print(type(primeiroDetalhamento), primeiroDetalhamento)
        print(type(segundoDetalhamento), segundoDetalhamento)
        print(type(P1), P1)
        print(type(P2), P2)
        print(type(P3), P3)
        print(type(P4), P4)
        print(type(P5), P5)
        print(type(P6), P6)

        # Atualiza a tabela de email_content
        conn = sql.connect('C:/Users/mauricio.longato/PycharmProjects/gmail_api/email_tracker.db')
        cur = conn.cursor()
        cur.executemany(
            '''INSERT INTO email_content(
                email_id, attachment_file_name, file_address, date_read
                , sentido_trade, tipo_ncm, modo_ncm, lista_ncm, bloco
                , pais, uf, porto, via, primeiroDetalhamento, segundoDetalhamento
                , P1, P2, P3, P4, P5, P6)
                VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);''',
                info_row)
        conn.commit()
        conn.close()

        # Atualiza status na tabela download inventory
        conn = sql.connect('C:/Users/mauricio.longato/PycharmProjects/gmail_api/email_tracker.db')
        conn.execute("update download_inventory set analysis_status=? where email_id=?", (1, email_id))
        conn.execute("update download_inventory set data_analysis=? where email_id=?",
                     (str(time.strftime("%d/%m/%Y %H:%M:%S")), email_id))
        conn.commit()
        conn.close()

    except:
        conn = sql.connect('C:/Users/mauricio.longato/PycharmProjects/gmail_api/email_tracker.db')
        conn.execute("update download_inventory set analysis_status=? where email_id=?", (3, email_id))
        conn.execute("update download_inventory set data_analysis=? where email_id=?",
                     (str(time.strftime("%d/%m/%Y %H:%M:%S")), email_id))
        conn.commit()
        conn.close()


