import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import sqlite3
import csv
from datetime import datetime

class SQLQueryManager:
    def __init__(self, root):
        self.root = root
        self.root.title("SQL Query Manager")
        self.root.geometry("1000x700")
        self.connection = None
        self.cursor = None
        
        # Configurar estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        self.create_widgets()
        
    def create_widgets(self):
        # Frame superior - Conexión
        connection_frame = ttk.LabelFrame(self.root, text="Conexión a Base de Datos", padding=10)
        connection_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(connection_frame, text="Ruta de BD:").grid(row=0, column=0, sticky="w", padx=5)
        self.db_path = ttk.Entry(connection_frame, width=50)
        self.db_path.grid(row=0, column=1, padx=5)
        self.db_path.insert(0, "database.db")
        
        ttk.Button(connection_frame, text="Conectar", command=self.connect_db).grid(row=0, column=2, padx=5)
        ttk.Button(connection_frame, text="Desconectar", command=self.disconnect_db).grid(row=0, column=3, padx=5)
        
        self.status_label = ttk.Label(connection_frame, text="Estado: Desconectado", foreground="red")
        self.status_label.grid(row=0, column=4, padx=20)
        
        # Frame medio - Consultas
        query_frame = ttk.LabelFrame(self.root, text="Consulta SQL", padding=10)
        query_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Área de texto para la consulta
        self.query_text = scrolledtext.ScrolledText(query_frame, height=8, width=80)
        self.query_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.query_text.insert("1.0", "SELECT * FROM tabla LIMIT 10;")
        
        # Botones de acción
        button_frame = ttk.Frame(query_frame)
        button_frame.pack(fill="x", pady=5)
        
        ttk.Button(button_frame, text="Ejecutar Consulta", command=self.execute_query).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Limpiar", command=self.clear_query).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Mostrar Tablas", command=self.show_tables).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Exportar CSV", command=self.export_csv).pack(side="left", padx=5)
        
        # Frame de resultados
        results_frame = ttk.LabelFrame(self.root, text="Resultados", padding=10)
        results_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Tabla de resultados con scrollbars
        table_container = ttk.Frame(results_frame)
        table_container.pack(fill="both", expand=True)
        
        # Scrollbars
        vsb = ttk.Scrollbar(table_container, orient="vertical")
        hsb = ttk.Scrollbar(table_container, orient="horizontal")
        
        self.results_tree = ttk.Treeview(table_container, 
                                         yscrollcommand=vsb.set, 
                                         xscrollcommand=hsb.set)
        vsb.config(command=self.results_tree.yview)
        hsb.config(command=self.results_tree.xview)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.results_tree.pack(fill="both", expand=True)
        
        # Label de información
        self.info_label = ttk.Label(results_frame, text="Registros: 0")
        self.info_label.pack(pady=5)
        
    def connect_db(self):
        try:
            db_path = self.db_path.get()
            self.connection = sqlite3.connect(db_path)
            self.cursor = self.connection.cursor()
            self.status_label.config(text="Estado: Conectado", foreground="green")
            messagebox.showinfo("Éxito", f"Conectado a {db_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al conectar: {str(e)}")
            
    def disconnect_db(self):
        if self.connection:
            self.connection.close()
            self.connection = None
            self.cursor = None
            self.status_label.config(text="Estado: Desconectado", foreground="red")
            messagebox.showinfo("Información", "Desconectado de la base de datos")
        else:
            messagebox.showwarning("Advertencia", "No hay conexión activa")
            
    def execute_query(self):
        if not self.connection:
            messagebox.showwarning("Advertencia", "Primero conecta a una base de datos")
            return
            
        query = self.query_text.get("1.0", tk.END).strip()
        if not query:
            messagebox.showwarning("Advertencia", "Ingresa una consulta SQL")
            return
            
        try:
            self.cursor.execute(query)
            
            # Si es un SELECT, mostrar resultados
            if query.strip().upper().startswith("SELECT"):
                results = self.cursor.fetchall()
                columns = [description[0] for description in self.cursor.description]
                self.display_results(columns, results)
            else:
                # Para INSERT, UPDATE, DELETE
                self.connection.commit()
                messagebox.showinfo("Éxito", f"Consulta ejecutada. Filas afectadas: {self.cursor.rowcount}")
                self.clear_results()
                
        except Exception as e:
            messagebox.showerror("Error", f"Error en la consulta: {str(e)}")
            
    def display_results(self, columns, results):
        # Limpiar resultados anteriores
        self.clear_results()
        
        # Configurar columnas
        self.results_tree["columns"] = columns
        self.results_tree["show"] = "headings"
        
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=150)
            
        # Insertar datos
        for row in results:
            self.results_tree.insert("", "end", values=row)
            
        self.info_label.config(text=f"Registros: {len(results)}")
        
    def clear_results(self):
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.results_tree["columns"] = []
        self.info_label.config(text="Registros: 0")
        
    def clear_query(self):
        self.query_text.delete("1.0", tk.END)
        
    def show_tables(self):
        if not self.connection:
            messagebox.showwarning("Advertencia", "Primero conecta a una base de datos")
            return
            
        try:
            self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = self.cursor.fetchall()
            
            if tables:
                table_list = "\n".join([table[0] for table in tables])
                messagebox.showinfo("Tablas en la BD", table_list)
            else:
                messagebox.showinfo("Información", "No hay tablas en la base de datos")
        except Exception as e:
            messagebox.showerror("Error", f"Error al obtener tablas: {str(e)}")
            
    def export_csv(self):
        items = self.results_tree.get_children()
        if not items:
            messagebox.showwarning("Advertencia", "No hay resultados para exportar")
            return
            
        try:
            filename = f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            columns = self.results_tree["columns"]
            
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(columns)
                
                for item in items:
                    values = self.results_tree.item(item)['values']
                    writer.writerow(values)
                    
            messagebox.showinfo("Éxito", f"Datos exportados a {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SQLQueryManager(root)
    root.mainloop()