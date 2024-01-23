import openpyxl
from PIL import Image, ImageDraw, ImageFont

font_light =ImageFont.truetype('Myriad Pro Semibold.ttf',50)
font_bold =ImageFont.truetype('MYRIADPRO-BOLD.OTF',65)
font_bold_numeration =ImageFont.truetype('MYRIADPRO-BOLD.OTF',40)

# Lendo excel
tableConsecrated = openpyxl.load_workbook('CONTROLE DIPLOMAS.xlsx');
pageYear = tableConsecrated['2023']



for indice,linha in enumerate(pageYear.iter_rows(min_row=2,max_row=2)):
    name = linha[0].value
    register = linha[1].value
    undergraduation = linha[2].value
    dataNow = linha[3].value
    keyUser = linha[4].value
    # Gerando certificado

    numeration = f'Numeração: {keyUser}'
    frase1 = f'A Associação Deo de Jiu-Jitsu, liderada pelo Grande Mestre Deoclécio Paulo - Faixa Vermelha 9 Grau, outorga a\n{name} ({register == name and " " or register}),\nna qualidade de faixa {undergraduation}, o presente diploma pelo grau de Mérito e Reconhecimento.' 

    image = Image.open('./CERTIFICADO BASE.jpg');
    desenhar = ImageDraw.Draw(image)
    desenhar.text((2650,540),numeration, fill='red', font=font_bold_numeration)
    desenhar.text((560,1220),frase1, fill='black', font=font_light, align="center",spacing=60)
    # desenhar.text((1150,1320),frase2, fill='black', align="center", font=font_bold)
    # desenhar.text((750,1420),frase3, fill='black', font=font_light)
    image.save(f'./certificado_final/{indice+1}-{name}.png')

# image = Image.open('./CERTIFICADO BASE.jpg');
# print(image.size)