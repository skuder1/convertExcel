import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import webbrowser


def abrir_linkedin():
    webbrowser.open("https://www.linkedin.com/in/pablo-passos-2ba525251/")


def atualizar_progresso(progress, valor):
    progress['value'] = valor
    app.update_idletasks()


def converter_arquivo():
    try:
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=(("Excel files", "*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv"), ("All files", "*.*"))
        )
        if not arquivo:
            return

        atualizar_progresso(progress, 20)

        if arquivo.endswith('.xls'):
            df = pd.read_excel(arquivo, engine='xlrd')
        elif arquivo.endswith('.xlsx') or arquivo.endswith('.xlsm'):
            df = pd.read_excel(arquivo, engine='openpyxl')
        elif arquivo.endswith('.xlsb'):
            df = pd.read_excel(arquivo, engine='pyxlsb')
        elif arquivo.endswith('.csv'):
            df = pd.read_csv(arquivo)
        else:
            messagebox.showerror("Erro", "Formato de arquivo não suportado!")
            return

        atualizar_progresso(progress, 60)

        formato_saida = formato_var.get()
        if formato_saida == "XLSX":
            arquivo_saida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel XLSX", "*.xlsx")])
            df.to_excel(arquivo_saida, index=False, engine='openpyxl')
        elif formato_saida == "XLS":
            arquivo_saida = filedialog.asksaveasfilename(defaultextension=".xls", filetypes=[("Excel XLS", "*.xls")])
            df.to_excel(arquivo_saida, index=False, engine='xlwt')
        elif formato_saida == "XLSM":
            arquivo_saida = filedialog.asksaveasfilename(defaultextension=".xlsm", filetypes=[("Excel XLSM", "*.xlsm")])
            df.to_excel(arquivo_saida, index=False, engine='openpyxl')
        elif formato_saida == "XLSB":
            messagebox.showinfo("Aviso", "Conversão para XLSB ainda não suportada diretamente pelo pandas.")
        elif formato_saida == "CSV":
            arquivo_saida = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
            df.to_csv(arquivo_saida, index=False)
        else:
            messagebox.showerror("Erro", "Formato de saída inválido!")

        atualizar_progresso(progress, 100)
        messagebox.showinfo("Sucesso", f"Arquivo convertido e salvo como {arquivo_saida}")
        atualizar_progresso(progress, 0)

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        atualizar_progresso(progress, 0)


app = tk.Tk()
app.title("Conversor de Arquivos Excel")
app.geometry("400x410")
app.config(bg="#f0f0f0")

titulo = tk.Label(app, text="Conversor de Arquivos Excel", font=("Helvetica", 16), bg="#f0f0f0", pady=10)
titulo.pack()

frame_formatos = tk.Frame(app, bg="#f0f0f0")
frame_formatos.pack(pady=10)

formato_label = tk.Label(frame_formatos, text="Escolha o formato de saída:", bg="#f0f0f0", font=("Helvetica", 12))
formato_label.grid(row=0, column=0, padx=10, pady=5)

formato_var = tk.StringVar(value="XLSX")
tk.Radiobutton(frame_formatos, text="XLSX", variable=formato_var, value="XLSX", bg="#f0f0f0").grid(row=1, column=0,
                                                                                                   sticky=tk.W, padx=20)
tk.Radiobutton(frame_formatos, text="XLS", variable=formato_var, value="XLS", bg="#f0f0f0").grid(row=2, column=0,
                                                                                                 sticky=tk.W, padx=20)
tk.Radiobutton(frame_formatos, text="XLSM", variable=formato_var, value="XLSM", bg="#f0f0f0").grid(row=3, column=0,
                                                                                                   sticky=tk.W, padx=20)
tk.Radiobutton(frame_formatos, text="XLSB", variable=formato_var, value="XLSB", bg="#f0f0f0").grid(row=4, column=0,
                                                                                                   sticky=tk.W, padx=20)
tk.Radiobutton(frame_formatos, text="CSV", variable=formato_var, value="CSV", bg="#f0f0f0").grid(row=5, column=0,
                                                                                                 sticky=tk.W, padx=20)

btn_converter = tk.Button(app, text="Converter Arquivo", command=converter_arquivo, bg="#4CAF50", fg="white",
                          font=("Helvetica", 12), padx=10, pady=5)
btn_converter.pack(pady=20)

progress = ttk.Progressbar(app, orient="horizontal", length=300, mode='determinate')
progress.pack(pady=10)

creditos_frame = tk.Frame(app, bg="#f0f0f0")
creditos_frame.pack(pady=20)

creditos_label = tk.Label(creditos_frame, text="Criado por", font=("Helvetica", 10), bg="#f0f0f0",
                          fg="#333333")
creditos_label.pack(side="left")

linkedin_label = tk.Label(creditos_frame, text="Pablo Passos", font=("Helvetica", 10, "underline"), bg="#f0f0f0", fg="blue",
                          cursor="hand2")
linkedin_label.pack(side="left", padx=5)
linkedin_label.bind("<Button-1>", lambda e: abrir_linkedin())

app.mainloop()