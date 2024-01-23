import openpyxl
from PIL import Image, ImageDraw, ImageFont
import textwrap
font_light =ImageFont.truetype('Myriad Pro Semibold Italic.ttf',50)
font_bold =ImageFont.truetype('Myriad Pro Bold Italic.ttf',85)
font_bold_numeration =ImageFont.truetype('Myriad Pro Semibold Italic.ttf',40)
# Lendo excel
tableConsecrated = openpyxl.load_workbook('CONTROLE DIPLOMAS.xlsx');
pageYear = tableConsecrated['2023']
for indice,linha in enumerate(pageYear.iter_rows(min_row=2, max_row=2)):
    name = linha[0].value
    register = linha[1].value
    undergraduation= linha[2].value
    dataOld = linha[3].value
    keyUser = linha[4].value
    data_certification = linha[5].value
    # Gerando certificado
    numberUser,data = keyUser.split('/')
    keyUserNew = f'{data}/{numberUser}'
    numeration = f'Numeração: {keyUserNew}'
    frase1 = f'A Associação Deo de Jiu-Jitsu, liderada pelo Grande Mestre Deoclécio Paulo - Faixa Vermelha 9º Grau, outorga a' 
    frase2 = f'{name} ({register == name and " " or register}),' 
    frase3 = f'na qualidade de faixa {undergraduation}, o presente diploma pelo grau de Mérito e Reconhecimento.' 
    dataEmissao = f'{data_certification}' 
    # Processamento de cada frase individualmente
    image = Image.open('./CERTIFICADO BASE.jpg')
    desenhar = ImageDraw.Draw(image)
    # Processamento de cada frase individualmente
    for i, frase in enumerate([frase1, frase2, frase3]):
            wrapped_text = textwrap.fill(frase, width=110)
            y_position = 1220 + i * 100  # Ajuste a posição vertical conforme necessário
            if i == 1 :  # Se for a frase 2, centralize horizontalmente
                text_bbox = desenhar.textbbox((0, 0), wrapped_text, font=font_bold)
                text_width = text_bbox[2] - text_bbox[0]
                x_position = (image.width - text_width) / 2
            elif i==2:
                text_bbox = desenhar.textbbox((0, 0), wrapped_text, font=font_light)
                text_width = text_bbox[2] - text_bbox[0]
                x_position = (image.width - text_width) / 2
            else:
                x_position = 560  # Posição normal para as outras frases
            font_to_use = font_bold if i == 1 else font_light
            desenhar.text((x_position, y_position), wrapped_text, fill='black',font=font_to_use, align="center", spacing=60)
    desenhar.text((2650, 540), numeration, fill='red', font=font_bold_numeration)
    desenhar.text((1350,1580), dataEmissao, fill='black', font=font_light)
    image.save(f'./certificado_final/{indice + 1}-{name}.png')