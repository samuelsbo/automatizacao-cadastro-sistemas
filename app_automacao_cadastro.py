# Esse projeto pode ser aplicado a qualquer programa, utilizei um sistema em site, mas pode ser aplicado a um programa no desktop
# como SAP, TOTVS (DataSul), entre outros sistemas.

# Esse código foi feito em um notebook com tela 15.6" e resolução 1920 x 1080, para rodar em telas com tamanho e resolução diferentes
# será necessário adaptar os valores de posição do mouse, você pode usar o mouseInfo para isso.

# Importando as bibliotecas necessárias ----------------------------------------------------------------------------------------------------------
import openpyxl # Ler arquivo excel
import pyautogui # Automatizar teclado e mouse
import pyperclip # Copiar textos com caractéres especiais
import time # Colocar "pausas" no código


# Mensagem de aviso para o usuário não mexer no mouse e no teclado enquanto o programa roda ------------------------------------------------------
pyautogui.alert("ATENÇÃO!\nO programa irá iniciar, não mexa em NADA enquanto o programa estiver rodando. Quando finalizar, irei te avisar.")


# Abrir o sistema --------------------------------------------------------------------------------------------------------------------------------
pyautogui.PAUSE = 0.7 # Adicionar uma pausa no código de 0.7 seg a cada ação do pyautogui
pyautogui.press('win') # Abertar o botão 'win'
pyautogui.write('chrome') # Digitar 'chrome' na barra de pesquisa do windowns
pyautogui.press('enter') # Abertar 'enter'
pyautogui.write('https://cadastro-produtos-devaprender.netlify.app/') # Digitar o site do sistema no chrome
pyautogui.press('enter') # Abertar 'enter'


# 01 - Importando a planilha em excel para o python ----------------------------------------------------------------------------------------------
workbook = openpyxl.load_workbook(r'produtos_ficticios.xlsx') # Carregar o arquivo excel do computador
sheet_produtos = workbook['Produtos'] # Selecionar a aba 'Produtos' da planilha


# Copiar a informação de um campo da planilha e colar no campo correspondente do sistema ---------------------------------------------------------
qtde_produtos_cadastrados = 0 # Variável para contar a quantidade de produtos que foram cadastrados pelo programa.

# Para cada linha da planilha o programa vai repetir o comportamento dentro do for, até que a primeira celula da linha esteja "vazia". 
for linha in sheet_produtos.iter_rows(min_row=2): # Os dados começam na 2ª linha da planilha

    if linha[0].value != '' and linha[0].value != None: # O conteúdo da primeira celula da linha tem que ser diferente de '' e None.
        qtde_produtos_cadastrados += 1 # Adicionar +1 na quantidade de produtos cadastrados
        
        # Nome do produto
        nome_produto = linha[0].value # "Selecionar" a celula da planilha
        pyperclip.copy(nome_produto) # Copiar o conteúdo da celula ("CTRL + C")
        pyautogui.click(220,277) # "Coordenada" do campo do sistema onde o robô tem que "clicar"
        pyautogui.hotkey('ctrl', 'v') # Colar o conteúdo da celula no campo.

        # Descrição
        descricao = linha[1].value
        pyperclip.copy(descricao)
        pyautogui.click(214,385)
        pyautogui.hotkey('ctrl', 'v')

        # Categoria
        categoria = linha[2].value
        pyperclip.copy(categoria)
        pyautogui.click(225,557)
        pyautogui.hotkey('ctrl', 'v')

        # codigo
        codigo = linha[3].value
        pyperclip.copy(codigo)
        pyautogui.click(222,664)
        pyautogui.hotkey('ctrl', 'v')

        # peso
        peso = linha[4].value
        pyperclip.copy(peso)
        pyautogui.click(215,772)
        pyautogui.hotkey('ctrl', 'v')

        # dimensoes
        dimensoes = linha[5].value
        pyperclip.copy(dimensoes)
        pyautogui.click(209,878)
        pyautogui.hotkey('ctrl', 'v')

        # "clicar em "PRÓXIMO" para avançar e continuar o processo.
        pyautogui.click(205,950)
        time.sleep(1) # Coloquei 1 seg de pausa no código para esperar o carregamento da página.

        # preco
        preco = linha[6].value
        pyperclip.copy(preco)
        pyautogui.click(196,310)
        pyautogui.hotkey('ctrl', 'v')

        # qtde_estoque
        qtde_estoque = linha[7].value
        pyperclip.copy(qtde_estoque)
        pyautogui.click(195,416)
        pyautogui.hotkey('ctrl', 'v')

        # data_validade
        data_validade = linha[8].value
        pyperclip.copy(data_validade)
        pyautogui.click(193,522)
        pyautogui.hotkey('ctrl', 'v')

        # cor
        cor = linha[9].value
        pyperclip.copy(cor)
        pyautogui.click(192,630)
        pyautogui.hotkey('ctrl', 'v')

        # tamanho
        pyautogui.click(209,734) #Clicar na opção de tamanho
        # CLica na opção de acordo com o tamanho
        tamanho = linha[10].value
        if tamanho == 'Pequeno':
            pyautogui.click(209,783)
        elif tamanho == 'Médio':
            pyautogui.click(208,818)
        else:
            pyautogui.click(215,855) #Se é diferente de pequeno e médio, então é grande.

        # material
        material = linha[11].value
        pyperclip.copy(material)
        pyautogui.click(192,844)
        pyautogui.hotkey('ctrl', 'v')

        # "clicar em "PRÓXIMO" para avançar e continuar o processo.
        pyautogui.click(194,917)
        time.sleep(1) # Coloquei 1 seg de pausa no código para esperar o carregamento da página.

        # fabricante
        fabricante = linha[12].value
        pyperclip.copy(fabricante)
        pyautogui.click(208,333)
        pyautogui.hotkey('ctrl', 'v')

        # pais_origem
        pais_origem = linha[13].value
        pyperclip.copy(pais_origem)
        pyautogui.click(218,439)
        pyautogui.hotkey('ctrl', 'v')

        # observacoes
        observacoes = linha[14].value
        pyperclip.copy(observacoes)
        pyautogui.click(222,539)
        pyautogui.hotkey('ctrl', 'v')

        # codigo de barras
        cod_barras = linha[15].value
        pyperclip.copy(cod_barras)
        pyautogui.click(197,712)
        pyautogui.hotkey('ctrl', 'v')

        # localizacao armazem
        loc_armazem = linha[16].value
        pyperclip.copy(loc_armazem)
        pyautogui.click(189,822)
        pyautogui.hotkey('ctrl', 'v')

        # "clicar em "CONCLUIR" para avançar e finalizar o cadastro.
        pyautogui.click(199,894)
        time.sleep(1) # Coloquei 1 seg de pausa no código para esperar o carregamento.

        # "clicar em "OK" para confirmar "Produto salvo no banco de dados!"
        pyautogui.click(1176,237)
        time.sleep(1) # Coloquei 1 seg de pausa no código para esperar o carregamento.

        # Clicar em "Adicionar novo item"
        pyautogui.click(970,618)
        time.sleep(1) # Coloquei 1 seg de pausa no código para esperar o carregamento.

        # Mensagem de aviso para o usuário informando que o programa finalizou e informando quantos produtos foram cadastrados--------------------
pyautogui.alert(f"PROGRAMA CONCLUÍDO COM SUCUESSO!\nForam cadastrados {qtde_produtos_cadastrados} produtos no sistema.")