import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import csv
import os

class SpreadsheetApp:
    def __init__(self, root, client_name):
        self.root = root
        self.client_name = client_name
        self.root.title(f"Planilha - {client_name}")
        self.data = []
        self.columns = ["A", "B", "C", "D", "E"]
        self.filepath = f"{client_name}.csv"
        self.load_data()
        
        style = ttk.Style()
        style.configure("Treeview", rowheight=25, font=("Arial", 12))
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
        style.layout("Treeview", [("Treeview.treearea", {'sticky': 'nswe'})])
        
        self.tree = ttk.Treeview(root, columns=self.columns, show='headings', style="Treeview")
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")
        
        self.tree.tag_configure("evenrow", background="#f0f0f0")
        self.tree.tag_configure("oddrow", background="#ffffff")
        
        self.tree.pack(expand=True, fill='both', padx=10, pady=10)
        
        self.add_row_button = tk.Button(root, text="Adicionar Linha", command=self.add_row, font=("Arial", 10, "bold"))
        self.add_row_button.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.save_button = tk.Button(root, text="Salvar", command=self.save_file, font=("Arial", 10, "bold"))
        self.save_button.pack(side=tk.RIGHT, padx=5, pady=5)
        
        self.export_button = tk.Button(root, text="Salvar Planilha Externa", command=self.export_file, font=("Arial", 10, "bold"))
        self.export_button.pack(side=tk.RIGHT, padx=5, pady=5)
        
        self.tree.bind('<Double-1>', self.edit_cell)
        self.refresh_tree()
    
    def load_data(self):
        if os.path.exists(self.filepath):
            with open(self.filepath, newline='') as file:
                reader = csv.reader(file)
                next(reader, None)  # Pular cabeçalho
                self.data = [row for row in reader]
    
    def add_row(self):
        self.data.append(["" for _ in self.columns])
        self.refresh_tree()
    
    def edit_cell(self, event):
        item = self.tree.selection()[0]
        col = self.tree.identify_column(event.x)[1:]
        col_idx = int(col) - 1
        
        entry_popup = tk.Toplevel(self.root)
        entry_popup.title("Editar Célula")
        
        entry = tk.Entry(entry_popup, font=("Arial", 12))
        entry.pack(padx=10, pady=10)
        entry.insert(0, self.tree.item(item, 'values')[col_idx])
        
        def save_edit():
            new_value = entry.get()
            row_idx = int(item) - 1
            self.data[row_idx][col_idx] = new_value
            self.refresh_tree()
            entry_popup.destroy()
        
        save_button = tk.Button(entry_popup, text="Salvar", command=save_edit, font=("Arial", 10, "bold"))
        save_button.pack(pady=10)
    
    def save_file(self):
        with open(self.filepath, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(self.columns)
            writer.writerows(self.data)
        messagebox.showinfo("Sucesso", "Dados salvos no aplicativo!")
    
    def export_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            with open(file_path, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(self.columns)
                writer.writerows(self.data)
            messagebox.showinfo("Sucesso", "Arquivo salvo para uso externo!")
    
    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for idx, row in enumerate(self.data):
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.tree.insert("", "end", iid=idx+1, values=row, tags=(tag,))

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lista de Clientes")
        
        self.clients = ["Cliente 1", "Cliente 2", "Cliente 3"]
        
        self.listbox = tk.Listbox(root, font=("Arial", 12))
        for client in self.clients:
            self.listbox.insert(tk.END, client)
        self.listbox.pack(padx=10, pady=10)
        
        self.show_button = tk.Button(root, text="Exibir Informações", command=self.show_spreadsheet, font=("Arial", 10, "bold"))
        self.show_button.pack(pady=10)
        
        self.add_client_button = tk.Button(root, text="Adicionar Cliente", command=self.add_client, font=("Arial", 10, "bold"))
        self.add_client_button.pack(pady=10)
        
        self.rename_client_button = tk.Button(root, text="Renomear Cliente", command=self.rename_client, font=("Arial", 10, "bold"))
        self.rename_client_button.pack(pady=10)
    
    def add_client(self):
        new_client = simpledialog.askstring("Novo Cliente", "Digite o nome do novo cliente:")
        if new_client:
            self.clients.append(new_client)
            self.listbox.insert(tk.END, new_client)
    
    def rename_client(self):
        selected_index = self.listbox.curselection()
        if not selected_index:
            messagebox.showwarning("Aviso", "Selecione um cliente primeiro!")
            return
        
        new_name = simpledialog.askstring("Renomear Cliente", "Digite o novo nome do cliente:")
        if new_name:
            self.clients[selected_index[0]] = new_name
            self.listbox.delete(selected_index)
            self.listbox.insert(selected_index, new_name)
    
    def show_spreadsheet(self):
        selected_index = self.listbox.curselection()
        if not selected_index:
            messagebox.showwarning("Aviso", "Selecione um cliente primeiro!")
            return
        
        client_name = self.listbox.get(selected_index)
        spreadsheet_window = tk.Toplevel(self.root)
        SpreadsheetApp(spreadsheet_window, client_name)

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
