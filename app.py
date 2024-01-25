import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2, max_row=2)):
    # Cada célula que contém a info que precisamos
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    # Transferir os dados da planilha para a imagem do certificado
    # Definir a fonte a ser usada
    font_nome = ImageFont.truetype('./tahomabd.ttf', size=90)  # Certifique-se de fornecer o tamanho correto
    # Ou
    font_geral = ImageFont.truetype('./tahoma.ttf', size=80)
    
    font_date = ImageFont.truetype('./tahoma.ttf', size=55)  # Certifique-se de fornecer o tamanho correto

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1020, 827), nome_participante, fill='black', font=font_nome)
    desenhar.text((1060, 950), nome_curso, fill='black', font=font_geral)
    desenhar.text((1480, 1182), str(carga_horaria), fill='black', font=font_geral)

    desenhar.text((750, 1770), data_inicio, fill='blue', font=font_date)
    desenhar.text((750, 1930), data_final, fill='blue', font=font_date)

    desenhar.text((2220, 1930), data_emissao, fill='blue', font=font_date)

    image.save(f'./{indice}_{nome_participante}_test.png')
