import os
import PySimpleGUI as sg

# Define o tema da interface
sg.theme('DarkAmber')

def carregar_projetos(path):
    projetos = []
    for file in os.listdir(path):
        if file.endswith('.pyw'):
            projetos.append(file)
    return projetos

import subprocess

def abrir_projeto(caminho):
    subprocess.run(['start', caminho], shell=True)


# Define o layout da primeira aba
tab1_layout = [[sg.Text('Projetos Salvos')],
               [sg.Listbox(values=carregar_projetos(os.path.join('empresa', 'Artefacts')), size=(30, 10), key='-PROJETOS-',
                           enable_events=True)],
               [sg.Button('Abrir Projeto', disabled=True, key='-ABRIR-')]]

# Define o layout da segunda aba
tab2_layout = [[sg.Text('Projetos Salvos')],
               [sg.Listbox(values=carregar_projetos(os.path.join('empresa', 'desenvolvimento', 'TabelaCadastro', 'src')),
                           size=(30, 10), key='-PROJETOS2-', enable_events=True)],
               [sg.Button('Abrir Projeto', disabled=True, key='-ABRIR2-')]]

# Cria o layout do TabGroup
layout = [[sg.TabGroup([
    [sg.Tab('Artefacts', tab1_layout), sg.Tab('Desenvolvimento', tab2_layout)]])],
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
            abrir_projeto(os.path.join('empresa', 'Artefacts', projeto_selecionado))
    if event == '-PROJETOS2-':
        if values['-PROJETOS2-']:
            window['-ABRIR2-'].update(disabled=False)
        else:
            window['-ABRIR2-'].update(disabled=True)
    if event == '-ABRIR2-':
        if values['-PROJETOS2-']:
            projeto_selecionado = values['-PROJETOS2-'][0]
            abrir_projeto(os.path.join('empresa', 'desenvolvimento', 'TabelaCadastro', 'src', projeto_selecionado))

window.close()
