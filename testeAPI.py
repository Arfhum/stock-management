import tkinter as tk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Autenticando e acessando/criando a planilha
def conectar_planilha():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credenciais.json", scope)
    client = gspread.authorize(creds)

    try:
        planilha = client.open("Estoque")
    except gspread.SpreadsheetNotFound:
        planilha = client.create("Estoque")
        planilha.share(None, perm_type='anyone', role='writer')  # opcional, deixa pública

    aba = planilha.sheet1
    if aba.row_count < 1 or aba.cell(1, 1).value != "Nome":
        aba.update('A1:C1', [["Nome", "Quantidade", "Validade"]])
    return aba

# Função para adicionar ou atualizar produto
def adicionar_ou_atualizar():
    nome = entrada_nome.get().strip()
    quantidade = entrada_qtd.get().strip()
    validade = entrada_validade.get().strip()

    if not nome or not quantidade or not validade:
        messagebox.showwarning("Aviso", "Preencha todos os campos!")
        return

    try:
        qtd = int(quantidade)
    except ValueError:
        messagebox.showerror("Erro", "Quantidade deve ser um número.")
        return

    aba = conectar_planilha()
    registros = aba.get_all_records()

    for i, linha in enumerate(registros, start=2):  # começa da linha 2
        if linha["Nome"].lower() == nome.lower():
            aba.update_cell(i, 2, quantidade)
            aba.update_cell(i, 3, validade)
            messagebox.showinfo("Atualizado", f"Produto '{nome}' atualizado com sucesso.")
            return

    # Se não encontrou, adiciona nova linha
    aba.append_row([nome, quantidade, validade])
    messagebox.showinfo("Adicionado", f"Produto '{nome}' adicionado com sucesso.")

# Interface com Tkinter
janela = tk.Tk()
janela.title("Gerenciador de Estoque (Google Planilhas)")

tk.Label(janela, text="Nome do Produto:").grid(row=0, column=0, sticky="e")
entrada_nome = tk.Entry(janela)
entrada_nome.grid(row=0, column=1)

tk.Label(janela, text="Quantidade:").grid(row=1, column=0, sticky="e")
entrada_qtd = tk.Entry(janela)
entrada_qtd.grid(row=1, column=1)

tk.Label(janela, text="Validade (AAAA-MM-DD):").grid(row=2, column=0, sticky="e")
entrada_validade = tk.Entry(janela)
entrada_validade.grid(row=2, column=1)

botao = tk.Button(janela, text="Adicionar / Atualizar", command=adicionar_ou_atualizar)
botao.grid(row=3, columnspan=2, pady=10)

janela.mainloop()
