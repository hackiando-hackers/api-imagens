import openpyxl
from PIL import Image

def matrix_to_image(matrix, image_path):
    width = len(matrix[0])
    height = len(matrix)

    img = Image.new("RGB", (width, height))

    for y, row in enumerate(matrix):
        for x, pixel in enumerate(row):
            r, g, b = pixel.strip("()").split(",")
            r, g, b = int(r), int(g), int(b)
            img.putpixel((x, y), (r, g, b))

    img.save(image_path)

def excel_to_matrix(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    matrix = []
    for row in ws.iter_rows():
        row_values = []
        for cell in row:
            value = cell.value
            row_values.append(value)
        matrix.append(row_values)

    return matrix

if __name__ == "__main__":
    excel_path = input("Digite o nome do arquivo Excel sem o formato: ") + ".xlsx"
    format_choice = int(input("Selecione o formato da imagem: 1 para PNG, 2 para JPG: "))

    if format_choice == 1:
        image_extension = ".png"
    elif format_choice == 2:
        image_extension = ".jpg"
    else:
        print("Opção inválida. Utilizando o formato PNG por padrão.")
        image_extension = ".png"

    image_restored_path = "image_restored" + image_extension

    print("Carregando. Aguarde...")
    matrix = excel_to_matrix(excel_path)
    matrix_to_image(matrix, image_restored_path)

    print(f"A imagem restaurada foi salva como '{image_restored_path}'.")