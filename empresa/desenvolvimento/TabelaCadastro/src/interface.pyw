import os
import PySimpleGUI as sg
from openpyxl import load_workbook


class TabelaCadastro:
    
    def __init__(self, arquivo):
        self.arquivo = arquivo
        self.planilha = 'TabelaCadastro'
        self.arquivo_aberto = load_workbook(self.arquivo)
        self.planilha_aberta = self.arquivo_aberto[self.planilha]

    # Define o tema da interface   
    sg.theme('DarkAmber')  
        
    def salvar_dados(self, id, barcode, quantidade):
        linha = (id, barcode, quantidade)
        self.planilha_aberta.append(linha)
        self.arquivo_aberto.save(self.arquivo)
        sg.Popup('Produto cadastrado com sucesso!', title='Sucesso')

      
class TelaCadastro:
    
    def __init__(self):
        caminho_tabela = os.path.join(os.getcwd(),"empresa","desenvolvimento", 'TabelaCadastro', 'data', 'Tabela.xlsx')
        self.tabela = TabelaCadastro(caminho_tabela)
        layout = [
            [sg.Text('ID', size=(15, 0)), sg.Input(size=(20, 0), key='id')],
            [sg.Text('Código de Barras', size=(15, 0)), sg.Input(size=(20, 0), key='barcode')],
            [sg.Text('Quantidade', size=(15, 0)), sg.Input(size=(20, 0), key='quantidade')],
            [sg.Button('Salvar')]
        ]
        self.janela = sg.Window('Cadastro de Produto').layout(layout)
        
        
    def Iniciar(self):
        while True:
            evento, valores = self.janela.Read()
            if evento == 'Salvar':
                id = valores['id']
                barcode = valores['barcode']
                quantidade = valores['quantidade']
                self.tabela.salvar_dados(id, barcode, quantidade)

            elif evento == sg.WIN_CLOSED:
                break
                
        self.janela.close()

tela = TelaCadastro()
tela.Iniciar() 
