from PySimpleGUI import PySimpleGUI as sg



sg.theme('DarkRed2')
layout =[
    [sg.Text('Departamento'),sg.Input(key='departamento')],
    [sg.Button('Acessar')]
]

janela = sg.Window('Sergecont Contabilidade', layout)

while True:
    eventos, valores = janela.read()
    if eventos == sg.WINDOW_CLOSED:
        break
    if eventos == 'Acessar':
        if valores['departamento'].lower() == 'fiscal':
            print("Seja Bem-vindo Fiscal")