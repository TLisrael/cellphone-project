import tkinter as tk
from tkinter import messagebox, ttk
import sqlite3
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

class CelularApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cadastro de Celulares")
        self.root.geometry("400x300")

        self.conn = sqlite3.connect('celulares.db')
        self.c = self.conn.cursor()

        self.c.execute('''CREATE TABLE IF NOT EXISTS celulares (
                            id INTEGER PRIMARY KEY,
                            modelo TEXT,
                            marca TEXT,
                            defeito INTEGER,
                            observacao TEXT
                            )''')
        self.conn.commit()

        self.container = tk.Frame(self.root)
        self.container.pack(pady=20)

        tk.Label(self.container, text="Modelo:", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10)
        self.entry_modelo = tk.Entry(self.container, font=("Arial", 12))
        self.entry_modelo.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.container, text="Marca:", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=10)
        self.entry_marca = tk.Entry(self.container, font=("Arial", 12))
        self.entry_marca.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(self.container, text="Defeito:", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=10)
        self.defeito_var = tk.BooleanVar()
        self.defeito_check = tk.Checkbutton(self.container, variable=self.defeito_var, text="Sim", font=("Arial", 12))
        self.defeito_check.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        tk.Label(self.container, text="Observação:", font=("Arial", 12)).grid(row=3, column=0, padx=10, pady=10)
        self.entry_observacao = tk.Entry(self.container, font=("Arial", 12))
        self.entry_observacao.grid(row=3, column=1, padx=10, pady=10)

        self.btn_cadastrar = tk.Button(self.container, text="Cadastrar", font=("Arial", 12), command=self.cadastrar_celular)
        self.btn_cadastrar.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

        self.btn_relatorio = tk.Button(self.container, text="Gerar Relatório Excel", font=("Arial", 12), command=self.gerar_relatorio)
        self.btn_relatorio.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

        self.btn_mostrar_cadastros = tk.Button(self.container, text="Mostrar Cadastros", font=("Arial", 12), command=self.mostrar_cadastros)
        self.btn_mostrar_cadastros.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

    def cadastrar_celular(self):
        modelo = self.entry_modelo.get()
        marca = self.entry_marca.get()
        defeito = 1 if self.defeito_var.get() else 0
        observacao = self.entry_observacao.get()

        self.c.execute("INSERT INTO celulares (modelo, marca, defeito, observacao) VALUES (?, ?, ?, ?)",
                       (modelo, marca, defeito, observacao))
        self.conn.commit()

        messagebox.showinfo("Cadastro", "Celular cadastrado com sucesso!")

        self.entry_modelo.delete(0, tk.END)
        self.entry_marca.delete(0, tk.END)
        self.defeito_check.deselect()
        self.entry_observacao.delete(0, tk.END)

    def gerar_relatorio(self):
        self.c.execute("SELECT modelo, marca, defeito, observacao FROM celulares")
        celulares = self.c.fetchall()

        df = pd.DataFrame(celulares, columns=['Modelo', 'Marca', 'Defeito', 'Observação'])

        file_path = "relatorio_celulares.xlsx"
        wb = Workbook()
        ws = wb.active

        # Converter o DataFrame em linhas e adicionar ao Excel
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        wb.save(file_path)

        messagebox.showinfo("Relatório", f"Relatório gerado com sucesso!\nSalvo em: {file_path}")

    def mostrar_cadastros(self):
        self.c.execute("SELECT * FROM celulares")
        celulares = self.c.fetchall()

        self.mostrar_cadastros_window = tk.Toplevel(self.root)
        self.mostrar_cadastros_window.title("Celulares Cadastrados")

        columns = ['ID', 'Modelo', 'Marca', 'Defeito', 'Observação']
        self.tree = ttk.Treeview(self.mostrar_cadastros_window, columns=columns, show='headings')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')

        for celular in celulares:
            self.tree.insert('', 'end', values=celular)

        self.tree.pack(padx=10, pady=10, fill='both', expand=True)

        self.mostrar_cadastros_window.protocol("WM_DELETE_WINDOW", self.fechar_mostrar_cadastros)

    def fechar_mostrar_cadastros(self):
        self.mostrar_cadastros_window.destroy()

    def __del__(self):
        self.conn.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = CelularApp(root)
    root.mainloop()
