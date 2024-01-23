import openpyxl
from PIL import Image, ImageDraw, ImageFont
import textwrap

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
    frase1 = f'A Associação Deo de Jiu-Jitsu, liderada pelo Grande Mestre Deoclécio Paulo - Faixa Vermelha 9 Grau, outorga a' 
    frase2 = f'{name} ({register == name and " " or register}),' 
    frase3 = f'na qualidade de faixa {undergraduation}, o presente diploma pelo grau de Mérito e Reconhecimento.' 
    # merged_text = f'A Associação Deo de Jiu-Jitsu, liderada pelo Grande Mestre Deoclécio Paulo - Faixa Vermelha 9 Grau, outorga a\n{name} ({register == name and " " or register}),\nna qualidade de faixa {undergraduation}, o presente diploma pelo grau de Mérito e Reconhecimento.' 
   
      # Processamento de cada frase individualmente
    image = Image.open('./CERTIFICADO BASE.jpg')
    desenhar = ImageDraw.Draw(image)

# Processamento de cada frase individualmente
for i, frase in enumerate([frase1, frase2, frase3]):
        wrapped_text = textwrap.fill(frase, width=110)
        
        y_position = 1220 + i * 100  # Ajuste a posição vertical conforme necessário
        if i == 1 or i==2:  # Se for a frase 2, centralize horizontalmente
            if(i==1):
                text_bbox = desenhar.textbbox((0, 0), wrapped_text, font=font_bold)
            else:
                text_bbox = desenhar.textbbox((0, 0), wrapped_text, font=font_light)
            text_width = text_bbox[2] - text_bbox[0]
            x_position = (image.width - text_width) / 2
        else:
            x_position = 560  # Posição normal para as outras frases
        desenhar.text((x_position, y_position), wrapped_text, fill='black', font=font_light, align="center", spacing=60)
    
        desenhar.text((2650, 540), numeration, fill='red', font=font_bold_numeration)
        image.save(f'./certificado_final/{indice + 1}-{name}.png')

# image = Image.open('./CERTIFICADO BASE.jpg');
# print(image.size)