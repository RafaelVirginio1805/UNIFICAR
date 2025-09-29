import os
import pandas as pd
from PyPDF2 import PdfMerger
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

# Pastas fixas
pasta_origem = r"C:/Downloads/Download_Boletos"
pasta_destino = r"C:/Desktop/materiais U.M/Unificar_arquivo/Boletos"
pasta_origem_nf = r"C:/Downloads/NF"
pasta_destino_nf = r"C:/Desktop/materiais U.M/Unificar_arquivo/NFs"

# Variáveis para planilhas
planilha_contratos = ""
planilha_nf = ""

os.makedirs(pasta_destino, exist_ok=True)
os.makedirs(pasta_destino_nf, exist_ok=True)


def processar_contratos(progress, logbox):
    global planilha_contratos
    if not planilha_contratos:
        messagebox.showerror("Erro", "Selecione a planilha de contratos.")
        return

    logbox.insert(END, "\n=== PROCESSANDO CONTRATOS ===\n")
    df = pd.read_excel(planilha_contratos)

    # Normaliza os contratos da planilha
    df['Contrato'] = df['Contrato'].astype(str).str.strip().str.upper()
    mapa_nomes = dict(zip(df['Contrato'], df['Cliente']))

    # Organiza os PDFs por contrato
    arquivos = {}
    for arquivo in os.listdir(pasta_origem):
        if arquivo.lower().endswith(".pdf") and "_" in arquivo:
            nome = os.path.splitext(arquivo)[0]
            contrato = nome.split("_")[0].strip().upper()
            arquivos.setdefault(contrato, []).append(os.path.join(pasta_origem, arquivo))

    total = len(df)
    progress["maximum"] = total
    progress["value"] = 0

    for i, contrato in enumerate(df['Contrato'], start=1):
        lista_pdfs = arquivos.get(contrato, [])

        if not lista_pdfs:
            logbox.insert(END, f"Contrato {contrato} não possui PDFs na pasta.\n")
            logbox.see(END)
            progress["value"] = i
            progress.update_idletasks()
            continue

        merger = PdfMerger()
        for pdf in sorted(lista_pdfs):
            merger.append(pdf)

        novo_nome = mapa_nomes.get(contrato, contrato) + ".pdf"
        caminho_saida = os.path.join(pasta_destino, novo_nome)
        merger.write(caminho_saida)
        merger.close()

        logbox.insert(END, f"Contrato {contrato} → {novo_nome}\n")
        logbox.see(END)

        progress["value"] = i
        progress.update_idletasks()

    messagebox.showinfo("Finalizado", "Boletos processados com sucesso!")


def processar_notas(progress, logbox):
    global planilha_nf
    if not planilha_nf:
        messagebox.showerror("Erro", "Selecione a planilha de notas fiscais.")
        return

    logbox.insert(END, "\n=== PROCESSANDO NOTAS FISCAIS ===\n")
    df = pd.read_excel(planilha_nf)

    mapa_nf_cliente = dict(zip(df['NF'].astype(str), df['Cliente']))
    arquivos_cliente, contador, intermediarios = {}, {}, []

    arquivos = [a for a in os.listdir(pasta_origem_nf) if a.lower().endswith(".pdf")]
    total = len(arquivos)
    progress["maximum"] = total
    progress["value"] = 0

    for i, arquivo in enumerate(arquivos, start=1):
        nf = os.path.splitext(arquivo)[0]
        cliente = mapa_nf_cliente.get(nf)

        if cliente:
            contador[cliente] = contador.get(cliente, 0) + 1
            sufixo = f"_{contador[cliente]:03d}"
            novo_nome = f"{cliente}{sufixo}.pdf"
            caminho_novo = os.path.join(pasta_destino_nf, novo_nome)

            merger = PdfMerger()
            merger.append(os.path.join(pasta_origem_nf, arquivo))
            merger.write(caminho_novo)
            merger.close()

            arquivos_cliente.setdefault(cliente, []).append(caminho_novo)
            intermediarios.append(caminho_novo)

            logbox.insert(END, f"NF {nf} → {novo_nome}\n")
        else:
            logbox.insert(END, f"NF {nf} não encontrada na planilha\n")

        logbox.see(END)
        progress["value"] = i
        progress.update_idletasks()

    # Mesclagem final
    for cliente, lista_pdfs in arquivos_cliente.items():
        if len(lista_pdfs) > 1:
            merger = PdfMerger()
            for pdf in sorted(lista_pdfs):
                merger.append(pdf)
            caminho_final = os.path.join(pasta_destino_nf, f"{cliente}.pdf")
            merger.write(caminho_final)
            merger.close()

            logbox.insert(END, f"Cliente {cliente} → {cliente}.pdf (mesclado {len(lista_pdfs)} NFs)\n")

            for arquivo in lista_pdfs:
                if os.path.exists(arquivo):
                    os.remove(arquivo)
        else:
            unico_pdf = lista_pdfs[0]
            novo_nome_final = os.path.join(pasta_destino_nf, f"{cliente}.pdf")
            os.rename(unico_pdf, novo_nome_final)

            logbox.insert(END, f"Cliente {cliente} com único arquivo renomeado.\n")
        logbox.see(END)

    messagebox.showinfo("Finalizado", "Notas fiscais processadas com sucesso!")


# Funções para seleção de planilha
def escolher_planilha_contratos():
    global planilha_contratos
    planilha_contratos = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_contratos.delete(0, END)
    entry_contratos.insert(0, planilha_contratos)


def escolher_planilha_nf():
    global planilha_nf
    planilha_nf = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_nf.delete(0, END)
    entry_nf.insert(0, planilha_nf)


def executar_tudo():
    if planilha_contratos:
        processar_contratos(progress, logbox)
    if planilha_nf:
        processar_notas(progress, logbox)
    if not planilha_contratos and not planilha_nf:
        messagebox.showwarning("Aviso", "Nenhuma planilha selecionada!")
    else:
        messagebox.showinfo("Finalizado", "Processamento completo concluído!")


# Interface gráfica
app = ttk.Window(themename="darkly")
app.title("Unificação de PDFs - Boletos e NFs")
app.geometry("850x650")

frame = ttk.Frame(app, padding=20)
frame.pack(fill=BOTH, expand=YES)

ttk.Label(frame, text="Planilha de Contratos (Boletos):").pack(anchor=W)
entry_contratos = ttk.Entry(frame, width=90)
entry_contratos.pack(anchor=W, pady=5)
ttk.Button(frame, text="Selecionar", command=escolher_planilha_contratos).pack(anchor=W, pady=5)

ttk.Label(frame, text="Planilha de Notas Fiscais:").pack(anchor=W)
entry_nf = ttk.Entry(frame, width=90)
entry_nf.pack(anchor=W, pady=5)
ttk.Button(frame, text="Selecionar", command=escolher_planilha_nf).pack(anchor=W, pady=5)

progress = ttk.Progressbar(frame, bootstyle=SUCCESS, length=650)
progress.pack(pady=15)

btns = ttk.Frame(frame)
btns.pack(pady=10)

ttk.Button(btns, text="Unificar Boletos", bootstyle=INFO, command=lambda: processar_contratos(progress, logbox)).grid(row=0, column=0, padx=10)
ttk.Button(btns, text="Unificar NFs", bootstyle=WARNING, command=lambda: processar_notas(progress, logbox)).grid(row=0, column=1, padx=10)
ttk.Button(btns, text="Unificar Tudo", bootstyle=PRIMARY, command=executar_tudo).grid(row=0, column=2, padx=10)

logbox = ttk.Text(frame, height=20)
logbox.pack(fill=BOTH, expand=YES, pady=10)

app.mainloop()
