import pandas as pd
from tkinter import Tk, Label, Entry, Button, messagebox, ttk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import pyautogui
import time

class FormPreenchimento:
    def __init__(self):
        self.dados = pd.read_excel("DadosFormulario.xlsx")
        self.janela = Tk()
        self.configurar_janela()
        self.criar_treeview()
        self.criar_campos_entrada()
        self.criar_botoes()
        self.janela.mainloop()

    def configurar_janela(self):
        self.janela.title("Preenchendo Formulários da WEB")
        estilo = ttk.Style()
        estilo.theme_use("alt")
        estilo.configure(".", font=("Arial", 20), rowheight=40)

    def criar_treeview(self):
        self.treeview = ttk.Treeview(self.janela, columns=(1, 2, 3, 4, 5), show="headings")
        colunas = ["Nome", "Email", "Sexo", "Estado", "Cor"]
        for i, coluna in enumerate(colunas, start=1):
            self.treeview.heading(str(i), text=coluna)
            self.treeview.column(str(i), anchor='center')

        self.treeview.grid(row=2, column=0, columnspan=10, sticky="NSEW")
        self.preencher_treeview()

    def preencher_treeview(self):
        for index in range(len(self.dados)):
            self.treeview.insert("", "end", values=tuple(self.dados.iloc[index]))

    def criar_campos_entrada(self):
        rotulos = ["Nome", "Email", "Sexo", "Estado", "Cor"]
        self.campos = []
        for i, rotulo in enumerate(rotulos):
            Label(self.janela, text=f"{rotulo}: ", font="Arial 12").grid(row=0, column=i*2, sticky="W")
            campo = Entry(self.janela, font="Arial 12")
            campo.grid(row=0, column=i*2 + 1, sticky="W")
            self.campos.append(campo)

    def criar_botoes(self):
        Button(self.janela, text="Preencher em Massa", font="Arial 20", command=self.preencher_em_massa).grid(row=1, column=0, columnspan=2, sticky="NSEW")
        Button(self.janela, text="Adicionar", font="Arial 20", command=self.adicionar_item).grid(row=1, column=2, columnspan=2, sticky="NSEW")
        Button(self.janela, text="Alterar", font="Arial 20", command=self.alterar_item).grid(row=1, column=4, columnspan=2, sticky="NSEW")
        Button(self.janela, text="Excluir", font="Arial 20", command=self.deletar_item).grid(row=1, column=6, columnspan=2, sticky="NSEW")
        Button(self.janela, text="Preenchimento Único", font="Arial 20", command=self.preencher_formulario_unico).grid(row=1, column=8, columnspan=2, sticky="NSEW")

        self.treeview.bind("<Double-1>", self.carregar_dados_campos)

    def preencher_em_massa(self):
        for item in self.treeview.get_children():
            linha = self.treeview.item(item)["values"]
            self.preencher_formulario_web(linha)

    def preencher_formulario_web(self, dados):
        with webdriver.Chrome(service=Service(ChromeDriverManager().install())) as navegador:
            navegador.get("https://pt.surveymonkey.com/r/79WHF9G")
            time.sleep(3)

            navegador.find_element(By.ID, "112904979").send_keys(dados[0])
            time.sleep(1)
            navegador.find_element(By.ID, "112904987").send_keys(dados[1])
            time.sleep(1)

            if dados[2] == "Masculino":
                navegador.find_element(By.ID, "112905004_855492046_label").click()
            else:
                navegador.find_element(By.ID, "112905004_855492047_label").click()
            time.sleep(2)

            self.selecionar_estado(navegador, dados[3])
            navegador.find_element(By.NAME, "112905214").send_keys(dados[4])
            time.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="patas"]/main/article/section/form/div[2]/button').click()
            time.sleep(2)

    def selecionar_estado(self, navegador, estado):
        dropdown = Select(navegador.find_element(By.ID, "112905103"))
        for option in dropdown.options:
            if option.text == estado:
                dropdown.select_by_visible_text(estado)
                break

    def adicionar_item(self):
        valores = [campo.get() for campo in self.campos]
        if any(not valor for valor in valores):
            messagebox.showinfo("Atenção!", "Preencha todos os campos")
            return

        self.treeview.insert("", "end", values=tuple(valores))
        messagebox.showinfo("Atenção!", "Dados adicionados com sucesso!")
        self.limpar_campos()

    def alterar_item(self):
        if not self.treeview.selection():
            messagebox.showinfo("Atenção!", "Selecione um item para alterar")
            return

        valores = [campo.get() for campo in self.campos]
        if any(not valor for valor in valores):
            messagebox.showinfo("Atenção!", "Preencha todos os campos")
            return

        item_selecionado = self.treeview.selection()[0]
        self.treeview.item(item_selecionado, values=tuple(valores))
        messagebox.showinfo("Atenção!", "Dados alterados com sucesso!")
        self.limpar_campos()

    def deletar_item(self):
        itens_selecionados = self.treeview.selection()
        for item in itens_selecionados:
            self.treeview.delete(item)
        messagebox.showinfo("Atenção!", "Item deletado com sucesso!")

    def preencher_formulario_unico(self):
        valores = [campo.get() for campo in self.campos]
        if any(not valor for valor in valores):
            messagebox.showinfo("Atenção!", "Preencha todos os campos")
            return

        self.preencher_formulario_web(valores)
        messagebox.showinfo("Atenção!", "Formulário preenchido com sucesso!")

    def carregar_dados_campos(self, event):
        item_selecionado = self.treeview.selection()[0]
        dados = self.treeview.item(item_selecionado, "values")
        for campo, valor in zip(self.campos, dados):
            campo.delete(0, "end")
            campo.insert(0, valor)

    def limpar_campos(self):
        for campo in self.campos:
            campo.delete(0, "end")

if __name__ == "__main__":
    FormPreenchimento()
