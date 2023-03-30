import os
from openpyxl import load_workbook
import PySimpleGUI as sg

# Define o tema da interface
sg.theme('DarkAmber') 

def abrir_script():
    script_path = os.path.join(os.getcwd(),"empresa","desenvolvimento", "TabelaCadastro", "src", "interface.pyw")
    os.system(f'start "" "{script_path}"')

def carregar_dados():
    file_path = os.path.join(os.getcwd(),"empresa","desenvolvimento", "TabelaCadastro", "data", "Tabela.xlsx")
    workbook = load_workbook(file_path)
    worksheet = workbook.active
    data = []
    for row in worksheet.iter_rows(values_only=True):
        data.append(row)
    return data
    
   

# Cria o layout da janela
layout = [[sg.Text('Tabela de Cadastro')],
          [sg.Table(values=carregar_dados(), headings=["ID", "Barcode", "Quantidade"], max_col_width=25,
                    auto_size_columns=False, display_row_numbers=False, justification='center', num_rows=20)],
          [sg.Button('Abrir Script')]]

# Cria a janela
window = sg.Window('Tabela de Cadastro', layout)

# Loop de eventos para interação com a janela
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == 'Abrir Script':
        abrir_script()

window.close()
