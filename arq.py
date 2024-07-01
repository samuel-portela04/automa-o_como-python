import openpyxl
import pyperclip
import pyautogui
from time import sleep

# entrar em planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_podutos = workbook['Produtos']
# copiar informação de um campo e colar no se campo correspondente
for linha in sheet_podutos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1103,356,duration=1)
    pyautogui.hotkey("ctrl","v")
    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1104,437,duration=1)
    pyautogui.hotkey("ctrl","v")

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1106,572,duration=1)
    pyautogui.hotkey("ctrl","v")

    codigo_do_produto = linha[3].value
    pyperclip.copy(codigo_do_produto)
    pyautogui.click(1105,665,duration=1)
    pyautogui.hotkey("ctrl","v")

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1103,749,duration=1)
    pyautogui.hotkey("ctrl","v")

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1106,836,duration=1)
    pyautogui.hotkey("ctrl","v")

    pyautogui.click(1126,891,duration=1)
    sleep(1)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1103,392,duration=1)
    pyautogui.hotkey("ctrl","v")

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1102,478,duration=1)
    pyautogui.hotkey("ctrl","v")

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1102,566,duration=1)
    pyautogui.hotkey("ctrl","v")

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1102,650,duration=1)
    pyautogui.hotkey("ctrl","v")

    tamanho = linha[10].value
    pyautogui.click(1129,734,duration=1)
    if tamanho == "Pequeno":
        pyautogui.click(1142,770,duration=1)
    elif tamanho == "Médio":
        pyautogui.click(1134,795,duration=1)
    else:
        pyautogui.click(1127,8170,duration=1)
    #ler info da planilha
    #se for "pequeno", clicar em uma posição
    #se for "médio", clicar em uma posição
    #se for "grande", clicar em uma posição
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1107,833,duration=1)
    pyautogui.hotkey("ctrl","v")

    pyautogui.click(1117,887,duration=1)
    sleep(1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1104,418,duration=1)
    pyautogui.hotkey("ctrl","v")

    pais_de_origem = linha[13].value
    pyperclip.copy(pais_de_origem)
    pyautogui.click(1108,511,duration=1)
    pyautogui.hotkey("ctrl","v")

    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(1112,582,duration=1)
    pyautogui.hotkey("ctrl","v")

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1114,726,duration=1)
    pyautogui.hotkey("ctrl","v")

    localizacao_de_armazem = linha[16].value
    pyperclip.copy(localizacao_de_armazem)
    pyautogui.click(1103,815,duration=1)
    pyautogui.hotkey("ctrl","v")
    #Botao de concluir
    pyautogui.click(1139,861,duration=1)
    #botao confirmar inclusao
    pyautogui.click(1610,193,duration=1)
    #botao confirmação 2
    pyautogui.click(1426,634,duration=1)


# repetir esses passos para outros campos até peencher cvampos daq
# clicar em próxima
# repetir os msmod passos e ir para a próxima página (página 2)
# repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em comcluir
# clicar em ok, para finalizar o processo
# clicar no ok mais uma vez na mensagem de confirmação de salvamento no banco de dados 
#clicar em adicionar mais um e repetir ate o processo ate finalizar a planilha

#PyAutoGUI(automação de clicks e teclado)
#Openpyxl Ileitura e automação de planilhas)