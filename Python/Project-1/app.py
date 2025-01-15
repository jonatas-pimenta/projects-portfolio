import PySimpleGUI as sg

sg.theme('reddit')

main_window = [
    [sg.Text('E-mail'), sg.Input(key='email')],
    [sg.Text('Senha'), sg.Input(key='senha', password_char='*')],
    [sg.FolderBrowse('Escolha Pasta Anexos', target='input_anexos'), sg.Input(key='input_anexos')],
    [sg.FolderBrowse('Escolher Pasta Planilha', target='input_planilhas'), sg.Input(key='input_planilhas')],
    [sg.Button('Salvar')]
]

window = sg.Window('Main', layout=main_window)

while True:
    events, values = window.read()
    if events == sg.WINDOW_CLOSED:
        break
    elif events == 'Salvar':
        email = values['email']
        senha = values['senha']
        caminho_pasta_anexos = values['input_anexos']
        caminho_pasta_planilhas = values['input_planilhas']
        print(f'O e-mail digitado foi {email}')
        print(f'A senha digitada foi {senha}')
        print(f'O caminho da pasta de anexos é {caminho_pasta_anexos}') 
        print(f'O caminho da pasta de planilhas é {caminho_pasta_planilhas}') 
          

