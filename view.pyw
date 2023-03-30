import os
from openpyxl import load_workbook
import PySimpleGUI as sg

# Define o tema da interface
sg.theme('DarkAmber') 

def carregar_projetos():
    projetos = []
    folder_path = os.path.join("empresa", "desenvolvimento", "TabelaCadastro", "src")
    for file in os.listdir(folder_path):
        if file.endswith('.pyw'):
            projetos.append(file)
    return projetos


def abrir_projeto(nome_projeto):
    projeto_path = os.path.join(os.getcwd(), "empresa", "desenvolvimento", "TabelaCadastro", "src", nome_projeto)
    os.system(f'start "" "{projeto_path}"')

# Cria o layout da janela
layout = [[sg.Text('Projetos Salvos')],
          [sg.Listbox(values=carregar_projetos(), size=(30, 10), key='-PROJETOS-', enable_events=True)],
          [sg.Button('Abrir Projeto', disabled=True, key='-ABRIR-')],
          [sg.Button('Fechar')]]

# Cria a janela
window = sg.Window('Projetos Salvos', layout)

# Loop de eventos para interação com a janela
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Fechar':
        break
    if event == '-PROJETOS-':
        if values['-PROJETOS-']:
            window['-ABRIR-'].update(disabled=False)
        else:
            window['-ABRIR-'].update(disabled=True)
    if event == '-ABRIR-':
        if values['-PROJETOS-']:
            projeto_selecionado = values['-PROJETOS-'][0]
            abrir_projeto(projeto_selecionado)

window.close()
