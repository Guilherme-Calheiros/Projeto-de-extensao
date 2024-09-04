import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

registros = []
preco_total = 0.0

precos_servico = {
    "Lavagem Simples": 25.00,
    "Lavagem Completa": 50.00,
    "Aspiração": 15.00,
    "Polimento": 100.00,
    "Higienização Interna": 100.00
}

def adicionar_registro():
    global preco_total
    cliente = entry_cliente.get()
    tipo_servico = selected_servico.get()

    if cliente and tipo_servico:
        preco = precos_servico[tipo_servico]
        preco_total += preco
        registros.append({
            'cliente': cliente,
            'tipo_servico': tipo_servico,
            'preco': preco,
            'horario': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        })
        messagebox.showinfo("Sucesso", f"Registro adicionado com sucesso!")
        entry_cliente.delete(0, tk.END)
        selected_servico.set(opcoes_servico[0])  
    else:
        messagebox.showwarning("Erro", "Por favor, preencha todos os campos.")

def gerar_planilha():
    global preco_total
    hoje = datetime.now().strftime('%Y-%m-%d')
    filename = f'registros_{hoje}.xlsx'

    wb = Workbook()
    ws = wb.active
    ws.title = "Registros"

    ws.append(['Nome do Cliente', 'Serviço prestado', 'Valor', 'Horário'])
    
    for registro in registros:
        ws.append([
            registro['cliente'],
            registro['tipo_servico'],
            f"R$ {registro['preco']:.2f}",
            registro['horario']
        ])
    
    ws.append([
        'TOTAL',
        '',
        f"R$ {preco_total:.2f}",
        ''
    ])
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    wb.save(filename)
    messagebox.showinfo("Planilha Gerada", f"Planilha salva como {filename}")
    registros.clear()
    preco_total = 0.0

root = tk.Tk()
root.title("Sistema de Registros - Lava-Jato")

tk.Label(root, text="Nome do Cliente:").grid(row=0, column=0, padx=10, pady=5)
entry_cliente = tk.Entry(root)
entry_cliente.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Tipo de Serviço:").grid(row=1, column=0, padx=10, pady=5)

opcoes_servico = list(precos_servico.keys())
selected_servico = tk.StringVar(value=opcoes_servico[0])

optionmenu_servico = tk.OptionMenu(root, selected_servico, *opcoes_servico)
optionmenu_servico.grid(row=1, column=1, padx=10, pady=5)

btn_adicionar = tk.Button(root, text="Adicionar Registro", command=adicionar_registro)
btn_adicionar.grid(row=3, column=0, padx=10, pady=10)

btn_planilha = tk.Button(root, text="Gerar Planilha", command=gerar_planilha)
btn_planilha.grid(row=3, column=1, padx=10, pady=10)

root.mainloop()