from tkinter import Tk, ttk, messagebox, filedialog
import xmltodict

from datetime import datetime
from collections import defaultdict

import json
import os


# Dados globais
dados_filtrados = []

def carregar_arquivos():
    global dados_filtrados
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos XML ou JSON",
        filetypes=[("Arquivos XML ou JSON", "*.xml *.json")]
    )
    if not arquivos:
        return

    total_itens = 0
    datas_venda = []
    produtos_por_dia = defaultdict(int)

    for caminho in arquivos:
        if caminho.endswith(".xml"):
            dados_dict = converter_xml_para_dict(caminho)
            if dados_dict:
                total_itens += extrair_dados(dados_dict, datas_venda, produtos_por_dia)
        elif caminho.endswith(".json"):
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    dados_dict = json.load(f)
                total_itens += extrair_dados(dados_dict, datas_venda, produtos_por_dia)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar {os.path.basename(caminho)}:\n{e}")
        try:
            os.remove(caminho.replace(".xml", ".json"))
        except:
            pass

    atualizar_tabela()

    if total_itens:
        dias = set(datas_venda)
        semanas = set(datetime.strptime(d, "%Y-%m-%d").isocalendar()[1] for d in datas_venda)
        meses = set(datetime.strptime(d, "%Y-%m-%d").strftime("%Y-%m") for d in datas_venda)
        msg = (
            f"{total_itens} item(s) carregado(s)\n"
            f"{len(produtos_por_dia)} dia(s) com vendas\n"
            f"Produtos vendidos por dia:\n" +
            "\n".join([f"{dia}: {qtde} item(s)" for dia, qtde in produtos_por_dia.items()]) +
            f"\n\nDistribuídos em {len(semanas)} semana(s) e {len(meses)} mês(es)"
        )
        messagebox.showinfo("Resumo de Vendas", msg)
    else:
        messagebox.showinfo("Nenhum item", "Nenhum dado de venda encontrado nos arquivos.")

def converter_xml_para_dict(caminho_xml):
    try:
        with open(caminho_xml, "r", encoding="utf-8") as f:
            xml_content = f.read()
        return xmltodict.parse(xml_content)
    except Exception as e:
        messagebox.showerror("Erro na conversão", f"Erro ao processar {os.path.basename(caminho_xml)}:\n{e}")
        return None

def extrair_dados(dados, datas_venda, produtos_por_dia):
    global dados_filtrados
    try:
        info = dados["nfeProc"]["NFe"]["infNFe"]
        det = info["det"]
        if isinstance(det, dict):
            det = [det]
        data_venda = info["ide"]["dhEmi"][:10]

        count = 0
        for item in det:
            prod = item["prod"]
            nome = prod.get("xProd", "Desconhecido")
            qtd = float(prod.get("qCom", 0))
            total = float(prod.get("vProd", 0))

            dados_filtrados.append({
                "Produto": nome,
                "Quantidade": qtd,
                "Valor Vendido": total,
                "Data da Venda": data_venda
            })
            datas_venda.append(data_venda)
            produtos_por_dia[data_venda] += 1
            count += 1
        return count
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao extrair dados:\n{e}")
        return 0

def atualizar_tabela(filtro=""):
    for row in tabela.get_children():
        tabela.delete(row)

    for item in dados_filtrados:
        if filtro.lower() in item["Produto"].lower() or filtro in item["Data da Venda"] or filtro in f"{item['Valor Vendido']:.2f}":
            tabela.insert("", "end", values=(
                item["Produto"],
                item["Quantidade"],
                f"R$ {item['Valor Vendido']:.2f}",
                item["Data da Venda"]
            ))

def pesquisar_produto():
    termo = entry_pesquisa.get().strip()
    atualizar_tabela(termo)

def mostrar_info_produto_selecionado(event=None):
    itens = tabela.selection()
    if not itens:
        return

    soma_qtd = 0
    soma_valor = 0
    for row in itens:
        valores = tabela.item(row)["values"]
        qtd = float(valores[1])
        valor = float(str(valores[2]).replace("R$", "").replace(",", "").strip())
        soma_qtd += qtd
        soma_valor += valor

    label_qtd_total.config(text=f"Total Quantidade: {soma_qtd}")
    label_valor_total.config(text=f"Total Vendido: R$ {soma_valor:.2f}")

# Interface
root = Tk()
root.title("Analisador de Vendas NFe - XML/JSON Multi Arquivo")
root.geometry("1120x680")
root.resizable(False, False)

frame = ttk.Frame(root, padding=10)
frame.pack(fill="both", expand=True)

ttk.Button(frame, text="Carregar XML ou JSON (vários arquivos)", command=carregar_arquivos).pack(pady=5)

ttk.Label(frame, text="Pesquisar Produto, Data ou Valor:").pack(pady=(10, 0))
entry_pesquisa = ttk.Entry(frame, width=50)
entry_pesquisa.pack()
ttk.Button(frame, text="Pesquisar", command=pesquisar_produto).pack(pady=5)

colunas = ("Produto", "Quantidade", "Valor Vendido", "Data da Venda")
tabela = ttk.Treeview(frame, columns=colunas, show="headings", height=20, selectmode="extended")
for col in colunas:
    tabela.heading(col, text=col)
    tabela.column(col, width=250 if col == "Produto" else 150, anchor="center")
tabela.pack(pady=10, fill="x")
tabela.bind("<<TreeviewSelect>>", mostrar_info_produto_selecionado)

info_frame = ttk.Frame(frame)
info_frame.pack(pady=5)

label_qtd_total = ttk.Label(info_frame, text="Total Quantidade: 0", font=("Arial", 12))
label_qtd_total.grid(row=0, column=0, padx=10)

label_valor_total = ttk.Label(info_frame, text="Total Vendido: R$ 0.00", font=("Arial", 12))
label_valor_total.grid(row=0, column=1, padx=10)

root.mainloop()
