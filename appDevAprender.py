import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('./dados_alunos/planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate (sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    fonte_nome = ImageFont.truetype('./fonts/tahomabd.ttf',90)
    fonte_geral = ImageFont.truetype('./fonts/tahoma.ttf',80)
    fonte_data = ImageFont.truetype('./fonts/tahoma.ttf',55)

    image = Image.open('./modelo_certifica/certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)
    desenhar.text((1050,827), nome_participante,fill='black',
    font=fonte_nome)
    desenhar.text((1080,954), nome_curso, fill='black',font=fonte_geral)
    desenhar.text((1440,1065), tipo_participacao, fill='black',font=fonte_geral)
    desenhar.text((1490,1189), f"{str(carga_horaria)} h", fill='black',font=fonte_geral)
    desenhar.text((750,1770), data_inicio, fill='blue',font=fonte_data)
    desenhar.text((750,1930), data_final, fill='blue',font=fonte_data)
    desenhar.text((2220,1930), data_emissao, fill='blue',font=fonte_data)

    image.save(f'./certificados/{indice} {nome_participante} certificado.png')

