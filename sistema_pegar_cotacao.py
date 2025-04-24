import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
from datetime import datetime
import requests
import pandas as pd
import numpy as np

def pegar_cotacao():
    moeda = comobox_selecionar_moeda.get()
    data = calendario_moeda.get()
    ano = data[-4:]
    mes= data[3:5]
    dia = data[:2]
    link = f'https://economia.awesomeapi.com.br/{moeda}-BRL/10?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    label_texto_cotacao['text'] = f'A cotação da moeda {moeda}, da data {data} é de R${valor_moeda}'

def selcionar_arquivo():
    caminho_arquivo_moeda = askopenfilename(title='Selecione o Arquivo de Moeda')
    var_arquivo.set(caminho_arquivo_moeda)
    if caminho_arquivo_moeda:
        label_arquivo_selecionado['text'] = f'Arquivo selecionado: {caminho_arquivo_moeda}'

def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_arquivo.get())
        moedas = df.iloc[:, 0]
        data_inicial = calendario_data_inical.get()
        data_final = calendario_data_final.get()
        ano_inicial = data_inicial[-4:]
        mes_inicial= data_inicial[3:5]
        dia_inicial = data_inicial[:2]
        ano_final = data_final[-4:]
        mes_final= data_final[3:5]
        dia_final = data_final[:2]
        for moeda in moedas:
            link = (f'https://economia.awesomeapi.com.br/{moeda}-BRL/10?'
                    f'start_date={ano_inicial}{mes_inicial}{dia_inicial}&'
                    f'end_date={ano_final}{mes_final}{dia_final}')
            requisicao_moeda = requests.get(link)
            cotacaos = requisicao_moeda.json()
            for cotacao in cotacaos:
                bid = float(cotacao['bid'])
                timestamp = int(cotacao['timestamp'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                if data not in df:
                    df[data] = np.nan
                df.loc[df.iloc[:, 0] == moeda, data] = bid
        df.to_excel('Teste.xlsx')
        label_atualizar_coatcoes['text'] = 'Arquivo atualizado com sucesso'
    except:
        label_atualizar_coatcoes['text'] = "Selecione um arquivo Excel no Formato Correto"

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')


dic_moedas = requisicao.json()

lista_moedas = list(dic_moedas.keys())

janela = tk.Tk()

janela.title('Ferramenta de Cotação de Moedas')

label_mensagem_cotacao = tk.Label(text='Cotação de 1 moeda específica', borderwidth=2, relief='solid')
label_mensagem_cotacao.grid(row=0, column=0, padx=10, pady=10, columnspan=3, sticky='nsew')

label_mensagem_selecionar_moedas = tk.Label(text='Selecione a moeda que desejar consultar:', anchor='e')
label_mensagem_selecionar_moedas.grid(row=1, column=0, padx=10, pady=10, columnspan=2, sticky='nsew')

comobox_selecionar_moeda = ttk.Combobox(values=lista_moedas)
comobox_selecionar_moeda.grid(row=1, column=2, padx=10, pady=10, sticky='nsew')

label_mensagem_selecionar_dia = tk.Label(text='Selecione o dia que deseja pegar a cotação:', anchor='e')
label_mensagem_selecionar_dia.grid(row=2, column=0, padx=10, pady=10, columnspan=2, sticky='nsew')

calendario_moeda = DateEntry(yaer=2025, locale='pt_br')
calendario_moeda.grid(row=2 , column=2, padx=10, pady=10, sticky='nsew')

label_texto_cotacao = tk.Label(text='')
label_texto_cotacao.grid(row=3, column=0, padx=10, pady=10, columnspan=2, sticky='nsew')

botao_pegar_cotacao = tk.Button(text='Pegar Cotação', command=pegar_cotacao)
botao_pegar_cotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')

label_mensagem_cotacoes = tk.Label(text='Cotação de Múltiplas Moedas', borderwidth=2, relief='solid')
label_mensagem_cotacoes.grid(row=4, column=0, padx=10, pady=10, columnspan=3, sticky='nsew')

label_selecionar_arquivo = tk.Label(text='Selecione um arquivo em Excel com as Moedas na Coluna A:', anchor='e')
label_selecionar_arquivo.grid(row=5, column=0, padx=10, pady=10, columnspan=2, sticky='nsew')

var_arquivo = tk.StringVar()

botao_selecionar_arquivo = tk.Button(text='Clique aqui para selecionar', command=selcionar_arquivo)
botao_selecionar_arquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

label_arquivo_selecionado = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
label_arquivo_selecionado.grid(row=6, column=0, padx=10, pady=10, columnspan=3, sticky='nsew')

label_data_inicial = tk.Label(text='Data Inicial', anchor='e')
label_data_inicial.grid(row=7, column=0, padx=10, pady=10, sticky='nsew')

calendario_data_inical = DateEntry(year=2025, locale='pt_br')
calendario_data_inical.grid(row=7, column= 1, padx=10, pady=10, sticky='nsew' )

label_data_final = tk.Label(text='Data Final', anchor='e')
label_data_final.grid(row=8, column=0, padx=10, pady=10, sticky='nsew')

calendario_data_final = DateEntry(year=2025, locale='pt_br')
calendario_data_final.grid(row=8, column=1, padx=10, pady=10, sticky='nsew')

botao_atualizar_cotacoes = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao_atualizar_cotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

label_atualizar_coatcoes = tk.Label(text='Arquivo de Moedas atualizado com sucesso.')
label_atualizar_coatcoes.grid(row=9, column=1, padx=10, pady=10, columnspan=2, sticky='nsew')

botao_fechar = tk.Button(text='Fechar', command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')

janela.mainloop()