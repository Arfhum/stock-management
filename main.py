import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import os

ARQUIVO_PLANILHA = "produtos.xlsx"

def salvar_produto():
    nome = entry_nome.get().strip()
    quantidade = entry_quantidade.get().strip()
    validade = entry_validade.get().strip()

    if not nome or not quantidade or not validade:
        messagebox.showwarning("Campos obrigatórios", "Preencha todos os campos.")
        return

    try:
        quantidade_int = int(quantidade)
    except ValueError:
        messagebox.showerror("Erro", "Quantidade deve ser um número inteiro.")
        return

    if os.path.exists(ARQUIVO_PLANILHA):
        try:
            planilha = load_workbook(ARQUIVO_PLANILHA)
            folha = planilha.active
        except InvalidFileException:
            messagebox.showerror("Erro", "Arquivo da planilha está corrompido ou inválido.")
            return
    else:
        planilha = Workbook()
        folha = planilha.active
        folha.append(["Nome", "Quantidade", "Validade"])  # Cabeçalhos

    folha.append([nome, quantidade_int, validade])
    planilha.save(ARQUIVO_PLANILHA)
    messagebox.showinfo("Sucesso", "Produto salvo com sucesso!")

    # Limpa os campos
    entry_nome.delete(0, tk.END)
    entry_quantidade.delete(0, tk.END)
    entry_validade.delete(0, tk.END)

def remover_quantidade_produto():
    nome = entry_nome.get().strip()
    quantidade = entry_quantidade.get().strip()

    if not nome or not quantidade:
        messagebox.showwarning("Campos obrigatórios", "Preencha o nome e a quantidade a ser removida.")
        return

    try:
        quantidade_a_remover = int(quantidade)
        if quantidade_a_remover <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Erro", "Quantidade inválida para remoção.")
        return

    if not os.path.exists(ARQUIVO_PLANILHA):
        messagebox.showerror("Erro", "Nenhuma planilha encontrada.")
        return

    try:
        planilha = load_workbook(ARQUIVO_PLANILHA)
        folha = planilha.active
    except InvalidFileException:
        messagebox.showerror("Erro", "Arquivo da planilha está corrompido ou inválido.")
        return

    linhas = list(folha.iter_rows(min_row=2))
    quantidade_restante = quantidade_a_remover
    alteracoes = False

    for row in linhas:
        nome_celula = row[0].value
        quantidade_celula = row[1].value

        if nome_celula and nome_celula.strip().lower() == nome.lower() and isinstance(quantidade_celula, int):
            if quantidade_celula >= quantidade_restante:
                row[1].value = quantidade_celula - quantidade_restante
                alteracoes = True
                break
            else:
                quantidade_restante = quantidade_celula
                row[1].value = 0
                alteracoes = True

    if not alteracoes:
        messagebox.showinfo("Não encontrado", f"Nenhum produto '{nome}' com estoque encontrado.")
    elif quantidade_restante > 0:
        messagebox.showwarning("Estoque insuficiente", f"Remoção parcial feita. Ainda faltam {quantidade_restante} unidade(s) para completar a remoção.")
    else:
        messagebox.showinfo("Sucesso", f"{quantidade_a_remover} unidade(s) de '{nome}' removida(s).")

    planilha.save(ARQUIVO_PLANILHA)

    # Limpa os campos
    entry_nome.delete(0, tk.END)
    entry_quantidade.delete(0, tk.END)
    entry_validade.delete(0, tk.END)

# Criação da interface gráfica
janela = tk.Tk()
janela.title("Cadastro de Produtos")
janela.geometry("300x320")
janela.resizable(False, False)

# Rótulos e campos de entrada
tk.Label(janela, text="Nome do Produto").pack(pady=5)
entry_nome = tk.Entry(janela, width=30)
entry_nome.pack()

tk.Label(janela, text="Quantidade").pack(pady=5)
entry_quantidade = tk.Entry(janela, width=30)
entry_quantidade.pack()

tk.Label(janela, text="Validade (ex: 12/2025)").pack(pady=5)
entry_validade = tk.Entry(janela, width=30)
entry_validade.pack()

# Botões
tk.Button(janela, text="Salvar Produto", command=salvar_produto).pack(pady=10)
tk.Button(janela, text="Remover Quantidade", command=remover_quantidade_produto).pack(pady=5)

janela.mainloop()
