import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    # Nome do produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1518, 325, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Descricao
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1282, 435, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1209, 545, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Codigo do produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1160, 631, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1166, 715, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1208, 804, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão próximo
    pyautogui.click(1127, 871, duration=1)
    sleep(3)

    # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1167, 368, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1123, 458, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1134, 535, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1147, 633, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Tamanho
    tamanho = linha[10].value
    pyautogui.click(1496, 730, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1149, 739, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1149, 765, duration=1)
    else:
        pyautogui.click(1149, 790, duration=1)

    # material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1106, 798, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    # Botão próximo
    pyautogui.click(1130, 858, duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1107, 385, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1114, 474, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1115, 564, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1111, 693, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1109, 778, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão concluir
    pyautogui.click(1129, 842, duration=1)
    # Botão confirmar inclusão
    pyautogui.click(1609, 190, duration=1)
    # Botão confirmação 2
    pyautogui.click(1432, 603, duration=1)
    # iniciar cadastro novamente
    pyautogui.click(1695, 580, duration=1)
