import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import random
import json
import pandas as pd
import openpyxl

class App:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Aplicativo de Sorteio")
        self.window.geometry("800x600")

        self.top_bar = tk.Frame(self.window)
        self.top_bar.pack(side=tk.TOP)

        self.import_button = tk.Button(self.top_bar, text="Importar", command=self.import_names)
        self.import_button.pack(side=tk.LEFT)

        self.clear_button = tk.Button(self.top_bar, text="Limpar", command=self.clear_sheet)
        self.clear_button.pack(side=tk.LEFT)

        self.export_format_var = tk.StringVar()
        self.export_format_var.set("-- SELECIONE --")

        self.export_dropdown = tk.OptionMenu(self.top_bar, self.export_format_var, "-- SELECIONE --", "JSON", "EXCEL")
        self.export_dropdown.pack(side=tk.LEFT)

        self.export_button = tk.Button(self.top_bar, text="Exportar", command=self.export_results)
        self.export_button.pack(side=tk.LEFT)

        self.exit_button = tk.Button(self.top_bar, text="Sair", command=self.window.quit, bg="red")
        self.exit_button.pack(side=tk.RIGHT)

        self.sheet_textbox = scrolledtext.ScrolledText(self.window, width=40, height=10)
        self.sheet_textbox.pack(fill=tk.BOTH, expand=True)

        self.bottom_bar = tk.Frame(self.window)
        self.bottom_bar.pack(side=tk.BOTTOM)

        self.summary_label = tk.Label(self.bottom_bar, text="Resumo das Importações")
        self.summary_label.pack()

        self.quantity_label = tk.Label(self.bottom_bar, text="Quantidade:")
        self.quantity_label.pack(side=tk.LEFT)

        self.quantity_entry = tk.Entry(self.bottom_bar)
        self.quantity_entry.pack(side=tk.LEFT)

        self.pick_button = tk.Button(self.bottom_bar, text="Sortear", command=self.pick_names, bg="yellow")
        self.pick_button.pack(side=tk.LEFT)

        self.results_textbox = scrolledtext.ScrolledText(self.window, width=40, height=10)
        self.results_textbox.pack(fill=tk.BOTH, expand=True)

        self.window.mainloop()

    def import_names(self):
        clipboard_text = self.window.clipboard_get()
        names = clipboard_text.split("\n")
        new_entries = 0
        repeated_entries = 0

        for name in names:
            name = name.strip()
            if name and not any(name in entry for entry in self.sheet_textbox.get("1.0", tk.END).split("\n")):
                new_entries += 1
                self.sheet_textbox.insert(tk.END, f"ID: {new_entries}\tNome: {name}\n")
            elif name:
                repeated_entries += 1

        self.summary_label.config(text=f"Novas entradas: {new_entries} | Entradas repetidas: {repeated_entries} | Total de nomes: {new_entries + repeated_entries}")

    def clear_sheet(self):
        self.sheet_textbox.delete('1.0', tk.END)
        self.summary_label.config(text="Resumo das Importações")

    def pick_names(self):
        quantity = self.quantity_entry.get()
        names_list = self.sheet_textbox.get("1.0", tk.END).split("\n")
        names_list = [name.split("\t")[1] for name in names_list if name]
        
        if quantity.isdigit() and int(quantity) <= len(names_list):
            results = random.sample(names_list, int(quantity))
            self.results_textbox.delete('1.0', tk.END)
            self.results_textbox.insert(tk.END, "\n".join(results))
        else:
            messagebox.showerror("Erro", "A quantidade inserida é inválida.")

    def export_results(self):
        export_format = self.export_format_var.get()

        if export_format == "JSON":
            self.export_to_json()
        elif export_format == "EXCEL":
            self.export_to_excel()
        else:
            messagebox.showerror("Erro", "Selecione um formato de exportação válido.")

    def export_to_json(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Arquivo JSON", "*.json")])

        if file_path:
            results = self.results_textbox.get("1.0", tk.END).splitlines()
            results_dict = {}

            for index, result in enumerate(results):
                results_dict[index + 1] = result

            with open(file_path, "w") as json_file:
                json.dump(results_dict, json_file, indent=4)

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilha Excel", "*.xlsx")])

        if file_path:
            results = self.results_textbox.get("1.0", tk.END).splitlines()
            results_list = []

            for result in results:
                results_list.append({"Resultado": result})

            df = pd.DataFrame(results_list)
            df.to_excel(file_path, index=False, engine="openpyxl")

app = App()
