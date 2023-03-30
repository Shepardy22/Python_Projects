from tkinter import *
import pyfirmata

# Conectando ao Arduino
board = pyfirmata.Arduino('COM3')

# Configurando o pino 13 como saída
pin13 = board.get_pin('d:13:o')

# Função para ligar/desligar o LED
def toggle_led():
    if pin13.read() == 0:
        pin13.write(1)
        label.config(text="LED ligado")
    else:
        pin13.write(0)
        label.config(text="LED desligado")

# Criando a interface gráfica
root = Tk()
root.title("Controle do LED")

# Adicionando o botão
button = Button(root, text="Ligar/Desligar LED", command=toggle_led)
button.pack(pady=10)

# Adicionando uma label para exibir o status
label = Label(root, text="LED desligado")
label.pack()

# Iniciando o loop principal da interface gráfica
root.mainloop()
