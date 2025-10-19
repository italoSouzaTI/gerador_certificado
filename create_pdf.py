import openpyxl
from PIL import Image, ImageDraw, ImageFont
import textwrap
import fitz  # PyMuPDF
import os
from datetime import datetime
import locale

# Configurar locale para PT-BR
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR')
    except:
        pass  # Se não conseguir, usar o formato padrão

# Carregando fontes Myriad Pro
# Caminhos possíveis das fontes (TTF ou OTF)
myriad_semibold_options = [
    os.path.expanduser('~/Library/Fonts/Myriad Pro Semibold Italic.ttf'),
    os.path.expanduser('~/Library/Fonts/myriad-pro-semibold-italic.otf'),
    os.path.expanduser('~/Library/Fonts/Myriad Pro Semibold Italic.otf'),
]
myriad_bold_options = [
    os.path.expanduser('~/Library/Fonts/Myriad Pro Bold Italic.ttf'),
    os.path.expanduser('~/Library/Fonts/myriad-pro-bold-italic.otf'),
    os.path.expanduser('~/Library/Fonts/Myriad Pro Bold Italic.otf'),
]

# Verificar quais fontes existem
myriad_semibold = None
myriad_bold = None

for path in myriad_semibold_options:
    if os.path.exists(path):
        myriad_semibold = path
        break

for path in myriad_bold_options:
    if os.path.exists(path):
        myriad_bold = path
        break

semibold_exists = myriad_semibold is not None
bold_exists = myriad_bold is not None

print(f"\n{'='*60}")
print("VERIFICANDO FONTES MYRIAD PRO:")
if semibold_exists:
    print(f"  Semibold Italic: ✓ ENCONTRADA")
    print(f"    Arquivo: {os.path.basename(myriad_semibold)}")
else:
    print(f"  Semibold Italic: ✗ NÃO ENCONTRADA")
    
if bold_exists:
    print(f"  Bold Italic: ✓ ENCONTRADA")
    print(f"    Arquivo: {os.path.basename(myriad_bold)}")
else:
    print(f"  Bold Italic: ✗ NÃO ENCONTRADA")

if not semibold_exists:
    print("\n⚠️  ATENÇÃO: Instale 'Myriad Pro Semibold Italic.ttf/otf' em ~/Library/Fonts/")
if not bold_exists:
    print("\n⚠️  ATENÇÃO: Instale 'Myriad Pro Bold Italic.ttf/otf' em ~/Library/Fonts/")
print(f"{'='*60}\n")

try:
    # Fontes para frase1 e frase3 (tamanho original 50 - 1 = 49)
    font_frase1_3 = ImageFont.truetype(myriad_semibold if semibold_exists else myriad_bold, 49)
    # Fonte para frase2 (tamanho original 85 - 1 = 84 bold)
    font_frase2 = ImageFont.truetype(myriad_bold, 84)
    # Fonte para numeração (tamanho original 40 - 1 = 39)
    font_bold_numeration = ImageFont.truetype(myriad_semibold if semibold_exists else myriad_bold, 39)
    # Fonte para graduação em negrito (tamanho 52 para destacar)
    font_undergrad_bold = ImageFont.truetype(myriad_bold, 52)
    # Fonte para data de emissão (tamanho original 50 - 1 = 49)
    font_data = ImageFont.truetype(myriad_semibold if semibold_exists else myriad_bold, 49)
    print("✓ Fontes Myriad Pro carregadas com sucesso!\n")
except Exception as e:
    print(f"⚠ Erro ao carregar Myriad Pro: {e}")
    print("Usando fontes do sistema como fallback...\n")
    try:
        # Tentando fontes do sistema macOS
        font_frase1_3 = ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial.ttf', 49)
        font_frase2 = ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial Bold.ttf', 84)
        font_bold_numeration = ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial.ttf', 39)
        font_undergrad_bold = ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial Bold.ttf', 52)
        font_data = ImageFont.truetype('/System/Library/Fonts/Supplemental/Arial.ttf', 49)
    except:
        font_frase1_3 = ImageFont.load_default()
        font_frase2 = ImageFont.load_default()
        font_bold_numeration = ImageFont.load_default()
        font_undergrad_bold = ImageFont.load_default()
        font_data = ImageFont.load_default()

# Convertendo PDF para imagem usando PyMuPDF
if not os.path.exists('./CERTIFICADO BASE.jpg'):
    print("Convertendo PDF para JPG...")
    pdf_document = fitz.open('CERTIFICADO BASE.pdf')
    page = pdf_document[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))  # 300 DPI
    pix.save('./CERTIFICADO BASE.jpg')
    pdf_document.close()

# Obter data atual em formato PT-BR
data_atual = datetime.now()
try:
    # Formato: "19 de outubro de 2025"
    data_certification = data_atual.strftime('%d de %B de %Y').upper()
except:
    # Fallback se locale não funcionar
    meses = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 
             'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
    data_certification = f'{data_atual.day} DE {meses[data_atual.month-1]} DE {data_atual.year}'

# Lendo excel
tableConsecrated = openpyxl.load_workbook('Controle de Diplomas.xlsx')
pageYear = tableConsecrated['2025']

# Contador de certificados gerados
certificados_gerados = 0

for indice,linha in enumerate(pageYear.iter_rows(min_row=2)):
    name = linha[0].value
    
    # Pular linhas vazias
    if not name:
        continue
        
    register = linha[1].value
    undergraduation= linha[2].value
    dataOld = linha[3].value
    keyUser = linha[5].value if len(linha) > 5 and linha[5].value else None
    
    # Debug: imprimir valores
    print(f"Nome: {name}, Registro: {register}, Graduação: {undergraduation}")
    print(f"KeyUser: {keyUser}")
    
    # Gerando certificado
    if keyUser and '/' in str(keyUser):
        numberUser, data = str(keyUser).split('/')
        keyUserNew = f'{data}/{numberUser}'
    else:
        # Se não houver keyUser, gerar um padrão
        keyUserNew = f'{indice+1}/2025'
    numeration = f'Numeração: {keyUserNew}'
    frase1 = f'A Associação Deo de Jiu-Jitsu, liderada pelo Grande Mestre Deoclécio Paulo - Faixa Vermelha 9º Grau, outorga a' 
    frase2 = f'{name} ({register == name and " " or register}),' 
    frase3_antes = f'na qualidade de faixa '
    frase3_depois = f', o presente diploma pelo grau de Mérito e Reconhecimento.'
    dataEmissao = f'BRASÍLIA - DF, {data_certification}' 
    
    # Processamento de cada frase individualmente
    image = Image.open('./CERTIFICADO BASE.jpg')
    desenhar = ImageDraw.Draw(image)
    
    # FRASE 1 (centralizada)
    wrapped_text1 = textwrap.fill(frase1, width=110)
    y_position1 = 1210
    text_bbox1 = desenhar.textbbox((0, 0), wrapped_text1, font=font_frase1_3)
    text_width1 = text_bbox1[2] - text_bbox1[0]
    x_position1 = (image.width - text_width1) / 2
    desenhar.text((x_position1, y_position1), wrapped_text1, fill='black', font=font_frase1_3, align="center", spacing=60)
    
    # FRASE 2 (Nome - centralizado e bold)
    wrapped_text2 = textwrap.fill(frase2, width=110)
    y_position2 = 1310
    text_bbox2 = desenhar.textbbox((0, 0), wrapped_text2, font=font_frase2)
    text_width2 = text_bbox2[2] - text_bbox2[0]
    x_position2 = (image.width - text_width2) / 2
    desenhar.text((x_position2, y_position2), wrapped_text2, fill='black', font=font_frase2, align="center", spacing=60)
    
    # FRASE 3 (com graduação em negrito)
    y_position3 = 1430
    # Montar a frase3 completa para centralizar
    frase3_completa = f'{frase3_antes}{undergraduation}{frase3_depois}'
    wrapped_text3 = textwrap.fill(frase3_completa, width=110)
    
    # Calcular posição X para centralizar
    text_bbox3 = desenhar.textbbox((0, 0), wrapped_text3, font=font_frase1_3)
    text_width3 = text_bbox3[2] - text_bbox3[0]
    x_position3 = (image.width - text_width3) / 2
    
    # Desenhar frase3 parte por parte com graduação em negrito
    x_atual = x_position3
    
    # Parte 1: "na qualidade de faixa "
    desenhar.text((x_atual, y_position3), frase3_antes, fill='black', font=font_frase1_3)
    bbox_antes = desenhar.textbbox((x_atual, y_position3), frase3_antes, font=font_frase1_3)
    x_atual = bbox_antes[2]
    
    # Parte 2: GRADUAÇÃO EM NEGRITO
    desenhar.text((x_atual, y_position3), undergraduation, fill='black', font=font_undergrad_bold)
    bbox_grad = desenhar.textbbox((x_atual, y_position3), undergraduation, font=font_undergrad_bold)
    x_atual = bbox_grad[2]
    
    # Parte 3: ", o presente diploma pelo grau de Mérito e Reconhecimento."
    desenhar.text((x_atual, y_position3), frase3_depois, fill='black', font=font_frase1_3)
    
    # Numeração e data
    desenhar.text((2650, 540), numeration, fill='red', font=font_bold_numeration)
    desenhar.text((1350, 1580), dataEmissao, fill='black', font=font_data)
    
    # Salvar certificado em PDF
    filename_pdf = f'./certificado_final/{indice + 1}-{name}.pdf'
    
    # Converter para RGB se necessário (PDF não suporta alguns modos de cor)
    if image.mode != 'RGB':
        image = image.convert('RGB')
    
    # Salvar como PDF
    image.save(filename_pdf, 'PDF', resolution=100.0, quality=95)
    certificados_gerados += 1
    print(f"✓ Certificado gerado: {filename_pdf}\n")

print(f"\n{'='*60}")
print(f"TOTAL DE CERTIFICADOS GERADOS: {certificados_gerados}")
print(f"{'='*60}")