import random
from PIL import Image
import os
from openpyxl import Workbook
import time
import colorsys

CORES = {
    "1": {"nome": "Vermelho", "valor": (255, 0, 0)},
    "2": {"nome": "Verde", "valor": (0, 255, 0)},
    "3": {"nome": "Azul", "valor": (0, 0, 255)},
    "4": {"nome": "Amarelo", "valor": (255, 255, 0)},
    "5": {"nome": "Magenta", "valor": (255, 0, 255)},
    "6": {"nome": "Ciano", "valor": (0, 255, 255)},
    "7": {"nome": "Marrom", "valor": (128, 0, 0)},
    "8": {"nome": "Verde Escuro", "valor": (0, 128, 0)},
    "9": {"nome": "Azul Escuro", "valor": (0, 0, 128)},
    "10": {"nome": "Oliva", "valor": (128, 128, 0)},
    "11": {"nome": "Roxo", "valor": (128, 0, 128)},
    "12": {"nome": "Teal", "valor": (0, 128, 128)},
    "13": {"nome": "Cinza", "valor": (128, 128, 128)},
    "14": {"nome": "Prata", "valor": (192, 192, 192)},
    "15": {"nome": "Branco", "valor": (255, 255, 255)},
    "16": {"nome": "Preto", "valor": (0, 0, 0)},
    "17": {"nome": "Laranja", "valor": (255, 165, 0)},
    "18": {"nome": "Rosa", "valor": (255, 192, 203)},
    "19": {"nome": "Turquesa", "valor": (64, 224, 208)},
    "20": {"nome": "Lavanda", "valor": (230, 230, 250)}
}



def carregar_imagem(caminho):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo '{caminho}' não encontrado.")
    
    try:
        imagem = Image.open(caminho)
        return imagem
    except Exception as e:
        print("Ocorreu um erro ao carregar a imagem:", e)

def imagem_para_matriz(imagem):
    largura, altura = imagem.size
    matriz = []

    for y in range(altura):
        linha = []
        for x in range(largura):
            pixel = imagem.getpixel((x, y))
            linha.append(pixel)
        matriz.append(linha)

    return matriz

def matriz_para_excel(matriz, caminho_imagem):
    wb = Workbook()
    ws = wb.active

    for y, linha in enumerate(matriz, start=1):
        for x, pixel in enumerate(linha, start=1):
            valor_pixel = f"({pixel[0]}, {pixel[1]}, {pixel[2]})"
            cell = ws.cell(row=y, column=x, value=valor_pixel)
            cell.alignment = cell.alignment.copy(wrapText=True)

    for coluna in ws.columns:
        max_length = 0
        coluna_letra = coluna[0].column_letter
        for cell in coluna:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[coluna_letra].width = adjusted_width

    nome_arquivo_original = os.path.splitext(os.path.basename(caminho_imagem))[0]
    nome_arquivo_convertido = nome_arquivo_original + '_convertido.xlsx'
    caminho_arquivo = os.path.join(os.path.dirname(__file__), nome_arquivo_convertido)
    wb.save(caminho_arquivo)

    print(f"Matriz salva com sucesso no arquivo: {caminho_arquivo}")

def verificar_cor_pixel(imagem, x, y):
    pixel = imagem.getpixel((x, y))
    print("Pixel selecionado:")
    print(f"A cor do pixel na posição ({x}, {y}) é: {pixel}")
    time.sleep(1)

    largura, altura = imagem.size

    pixels_mesma_cor = []
    for py in range(altura):
        for px in range(largura):
            if imagem.getpixel((px, py)) == pixel:
                pixels_mesma_cor.append((px, py))

    print("Número de pixels com a mesma cor: Selecionados!")
    time.sleep(1)

    return pixel, pixels_mesma_cor

def alterar_cor_pixels(imagem, pixels, nova_cor):
    for x, y in pixels:
        imagem.putpixel((x, y), nova_cor)





def rgb_to_cmyk(rgb):
    r, g, b = rgb
    c = 1 - (r / 255)
    m = 1 - (g / 255)
    y = 1 - (b / 255)
    k = min(c, m, y)
    if k == 1:
        return 0, 0, 0, 1
    return (c - k) / (1 - k), (m - k) / (1 - k), (y - k) / (1 - k), k





def rgb_para_escala_de_cinza_media_ponderada(rgb):
    r, g, b = rgb
    y = 0.299 * r + 0.587 * g + 0.114 * b
    return int(y), int(y), int(y)

def rgb_para_escala_de_cinza_luminosidade(rgb):
    r, g, b = rgb
    y = 0.21 * r + 0.72 * g + 0.07 * b
    return int(y), int(y), int(y)

def rgb_para_escala_de_cinza_dessaturacao(rgb):
    r, g, b = rgb
    y = (r + g + b) / 3
    return int(y), int(y), int(y)

def rgb_para_escala_de_cinza_maximo(rgb):
    max_value = max(rgb)
    return max_value, max_value, max_value

def rgb_para_escala_de_cinza_minimo(rgb):
    min_value = min(rgb)
    return min_value, min_value, min_value

def converter_para_escala_de_cinza(imagem, metodo):
    largura, altura = imagem.size
    for y in range(altura):
        for x in range(largura):
            pixel = imagem.getpixel((x, y))
            nova_cor = metodo(pixel)
            imagem.putpixel((x, y), nova_cor)

def converter_todos_para_escala_de_cinza(imagem, nome_arquivo):
    # Métodos de conversão de escala de cinza
    metodos = {
        "Média Ponderada": rgb_para_escala_de_cinza_media_ponderada,
        "Luminosidade": rgb_para_escala_de_cinza_luminosidade,
        "Dessaturação": rgb_para_escala_de_cinza_dessaturacao,
        "Decomposição de Cores (Máximo)": rgb_para_escala_de_cinza_maximo,
        "Decomposição de Cores (Mínimo)": rgb_para_escala_de_cinza_minimo
    }

    for nome_metodo, metodo in metodos.items():
        imagem_temp = imagem.copy()
        converter_para_escala_de_cinza(imagem_temp, metodo) 
        novo_nome_arquivo = nome_arquivo + f"_convertido_{nome_metodo.replace(' ', '_').lower()}.png"
        imagem_temp.save(novo_nome_arquivo)
        print(f"Imagem convertida para escala de cinza ({nome_metodo}) e salva como: {novo_nome_arquivo}")

def main():
    imagem = None 

    print("Imagens .png & .jpg são aceitas.")
    time.sleep(1.5)
    opcao_parte1 = input("Deseja executar a Parte 1 (Conversão da imagem para matriz no Excel)? (s/n): ").lower()
    if opcao_parte1 == 's':
        print("PARTE 1: IMAGEM PARA MATRIZ NO EXCEL")
        time.sleep(1.5)
        nome_arquivo = input("Digite o nome do arquivo de imagem (sem o formato): ")
        
        caminho_png = nome_arquivo + ".png"
        caminho_jpg = nome_arquivo + ".jpg"

        if os.path.exists(caminho_png):
            caminho_imagem = caminho_png
        elif os.path.exists(caminho_jpg):
            caminho_imagem = caminho_jpg
        else:
            print("Arquivo não encontrado.")
            return
        
        imagem = carregar_imagem(caminho_imagem)
        
        if imagem:
            print("Imagem carregada com sucesso!")
            matriz_pixels = imagem_para_matriz(imagem)
            print("Matriz de pixels criada.")
            print("Dimensões da matriz:", len(matriz_pixels), "x", len(matriz_pixels[0]))

            matriz_para_excel(matriz_pixels, caminho_imagem)
            time.sleep(1.5)
    elif opcao_parte1 == 'n':
        print("Você optou por pular a Parte 1.")
        time.sleep(1.5)
    else:
        print("Opção inválida.")
        return

    if opcao_parte1 == 'n' or opcao_parte1 == 's':  
        opcao_parte2 = input("Deseja executar a Parte 2 (Modificar cores da imagem)? (s/n): ").lower()
        time.sleep(1.5)
        if opcao_parte2 == 's':
            print("PARTE 2: CONVERTER CORES DA IMAGEM")
            if not imagem:  
                nome_arquivo = input("Digite o nome do arquivo de imagem (sem o formato): ")
                caminho_png = nome_arquivo + ".png"
                caminho_jpg = nome_arquivo + ".jpg"

                if os.path.exists(caminho_png):
                    caminho_imagem = caminho_png
                elif os.path.exists(caminho_jpg):
                    caminho_imagem = caminho_jpg
                else:
                    print("Arquivo não encontrado.")
                    return
                
                imagem = carregar_imagem(caminho_imagem)
                if not imagem:
                    return
            
            opcao = input("Escolha uma opção:\n1. Digitar coordenadas do pixel\n2. Coordenadas aleatórias\nOpção: ")

            if opcao == "1":
                x_pixel = int(input("Digite a coordenada X do pixel que deseja verificar: "))
                y_pixel = int(input("Digite a coordenada Y do pixel que deseja verificar: "))
                cor_pixel, pixels_mesma_cor = verificar_cor_pixel(imagem, x_pixel, y_pixel)
            elif opcao == "2":
                largura, altura = imagem.size
                x_pixel = random.randint(0, largura - 1)
                y_pixel = random.randint(0, altura - 1)
                print("Pixel selecionado:")
                print(f"Coordenadas do pixel selecionado: ({x_pixel}, {y_pixel})")
                time.sleep(1)
                cor_pixel, pixels_mesma_cor = verificar_cor_pixel(imagem, x_pixel, y_pixel)
            else:
                print("Opção inválida.")
                return

            print("Selecione uma nova cor:")
            for chave, cor_info in CORES.items():
                print(f"{chave}: {cor_info['nome']} - {cor_info['valor']}")
            escolha_cor = input("Opção: ")

            nova_cor = CORES.get(escolha_cor)
            if nova_cor is None:
                print("Opção de cor inválida.")
                return

            alterar_cor_pixels(imagem, pixels_mesma_cor, nova_cor["valor"])
            
            novo_nome_arquivo = nome_arquivo + "_modificado.png"
            imagem.save(novo_nome_arquivo)
            print(f"Imagem modificada salva como: {novo_nome_arquivo}")
        elif opcao_parte2 == 'n':
            print("Você optou por pular a Parte 2.")
            time.sleep(1.5)
        else:
            print("Opção inválida.")
            return

    opcao_parte3 = input("Deseja executar a Parte 3 (Converter imagem para CMYK ou escala de cinza)? (s/n): ").lower()
    print("PARTE 3: CONVERSÃO DE RGB PARA CMYK OU ESCALA DE CINZA")
    time.sleep(1.5)
    if opcao_parte3 == 's':
        if not imagem:  
            nome_arquivo = input("Digite o nome do arquivo de imagem (sem o formato): ")
            caminho_png = nome_arquivo + ".png"
            caminho_jpg = nome_arquivo + ".jpg"

            if os.path.exists(caminho_png):
                caminho_imagem = caminho_png
            elif os.path.exists(caminho_jpg):
                caminho_imagem = caminho_jpg
            else:
                print("Arquivo não encontrado.")
                return
            
            imagem = carregar_imagem(caminho_imagem)
            if not imagem:
                return

        opcao = input("Escolha uma opção:\n1. Converter para CMYK\n2. Converter para Escala de Cinza\nOpção: ")

        if opcao == "1":
            # Código de conversão para CMYK
            img = Image.open(caminho_imagem)
            if img.mode != 'RGB':
                img = img.convert('RGB')

            cmyk_img = Image.new('CMYK', img.size)
            pixels = cmyk_img.load()

            for i in range(img.size[0]):
                for j in range(img.size[1]):
                    rgb = img.getpixel((i, j))
                    cmyk = rgb_to_cmyk(rgb)
                    pixels[i, j] = tuple(int(255 * x) for x in cmyk)

            cmyk_img.save("output_cmyk.jpg")
            print("imagem convertida com sucesso para CMYK")

        elif opcao == "2":
            # Código de conversão para escala de cinza...
                converter_todos_para_escala_de_cinza(imagem, nome_arquivo)
                return
    elif opcao_parte3 == 'n':
        print("Você optou por pular a Parte 3.")
        time.sleep(1.5)
    else:
        print("Opção inválida.")

if __name__ == "__main__":
    main()
