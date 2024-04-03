from math import sqrt
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
    #fontes usados no projeto
    fonte_nome = ImageFont.truetype('./fonts/tahomabd.ttf',50)
    fonte_geral = ImageFont.truetype('./fonts/tahoma.ttf',35)
    fonte_data = ImageFont.truetype('./fonts/tahoma.ttf',30)
    #carregamento do modelo de certificado
    image = Image.open('./modelo_certifica/modelo_cert.png')
    desenhar = ImageDraw.Draw(image)
    #desenhar nome do participante
    largura_imagem, altura_imagem = image.size
    largura_texto = len(nome_participante)
    posicao_x_central = (largura_imagem - largura_texto)//2
    # essa forma que encontrei pra centralizar o texto na imagem
    desenhar.text((posicao_x_central-(largura_texto*12),400), nome_participante,fill='green', font=fonte_nome)
    # desenhar mome do curso
    desenhar.text((590,500), nome_curso, fill='blue',font=fonte_geral)
    # desenhar tipo de participação
    desenhar.text((430,550), tipo_participacao, fill='blue',font=fonte_geral)
    # desenhar carga horária
    desenhar.text((900,550), f"{str(carga_horaria)} h", fill='blue',font=fonte_geral)
    # desenhar a data inicial
    desenhar.text((300,600), data_inicio, fill='blue',font=fonte_data)
    # desenhar a data final 
    desenhar.text((732,600), data_final, fill='blue',font=fonte_data)
    #salvando as imagens
    image.save(f'./certificados/{indice} {nome_participante} certificado.png')
    

   
    
