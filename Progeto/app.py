import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

class CadastroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cadastro App")

        # Botão para carregar o arquivo Excel
        btn_carregar = tk.Button(root, text="Carregar Excel", command=self.carregar_excel)
        btn_carregar.pack()

        # Botão para cadastrar dados
        btn_cadastrar = tk.Button(root, text="Cadastrar", command=self.cadastrar)
        btn_cadastrar.pack()

        # Botão para excluir dados
        btn_excluir = tk.Button(root, text="Excluir", command=self.excluir)
        btn_excluir.pack()

    def carregar_excel(self):
        # Abrir a caixa de diálogo para selecionar o arquivo Excel
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])

        # Ler o arquivo Excel
        try:
            self.df = pd.read_excel(file_path)
            messagebox.showinfo("Sucesso", "Arquivo Excel carregado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo Excel: {str(e)}")

    def cadastrar(self):
        # Lógica para cadastrar dados no Excel
        # Se o arquivo não existir, cria um DataFrame vazio
        # Aqui você deve usar a biblioteca pandas para adicionar dados ao DataFrame self.df
        # Exemplo:
            novo_dado = {'Nome': 'João', 'Email': 'joao@email.com', 'Centro de Custo': 'CC001', 'Valor': 100, 'Setor': 'RH'}
            self.df = self.df.append(novo_dado, ignore_index=True)

            self.df.to_excel("Rateio atual eMAIL - modificado.xlsx", index=False)

            messagebox.showinfo("Sucesso","Dados cadastrados com sucesso!")

        #messagebox.showinfo("Erro", "Carregue um arquivo Excel !")

    def excluir(self):
        # Lógica para excluir dados no Excel
        # Aqui você deve usar a biblioteca pandas para remover dados do DataFrame self.df
        # Exemplo:
        # self.df = self.df[self.df['Email'] != 'joao@email.com']
        # self.df.to_excel("novo_arquivo.xlsx", index=False)
        # Verificar se o arquivo Excel foi carregado 
        if self.df is None:
            messagebox.showerror("Erro", "Carregue um arquivo Excel !")
            return

            # Solicita ao usuário o ID do registro a ser excluído
        try:
            id_registro = int(input("Informe o ID do registro a ser excluído: "))
        except ValueError:
            messagebox.showerror("Erro", "O ID do registro deve ser um número inteiro!")
            return
        messagebox.showinfo("Sucesso", "Dados excluídos com sucesso!")


if __name__ == "__main__":
    root = tk.Tk()
    app = CadastroApp(root)
    root.mainloop()
