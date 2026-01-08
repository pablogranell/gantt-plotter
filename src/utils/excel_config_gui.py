import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd


class ExcelConfigGUI:
    def __init__(self, on_load=None):
        self.on_load = on_load
        self.file_path = None
        self.sheets = []
        self.columns = []
        
        self.root = tk.Tk()
        self.root.title("Configuracion de Gantt")
        self.root.geometry("375x425")
        self.root.resizable(True, True)
        
        self._create_widgets()
        
    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Archivo
        file_frame = ttk.LabelFrame(main_frame, text="Archivo Excel", padding="5")
        file_frame.pack(fill=tk.X, pady=5)
        
        self.file_label = ttk.Label(file_frame, text="Ninguno seleccionado")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(file_frame, text="Seleccionar", command=self._select_file).pack(side=tk.RIGHT)
        
        # Ajustes de hoja (Hoja + Cabecera juntos)
        sheet_frame = ttk.LabelFrame(main_frame, text="Ajustes de hoja", padding="5")
        sheet_frame.pack(fill=tk.X, pady=5)
        
        # Contenedor izquierdo: Hoja
        left_container = ttk.Frame(sheet_frame)
        left_container.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(left_container, text="Hoja:").pack(side=tk.LEFT)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(left_container, textvariable=self.sheet_var, state="readonly", width=20)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_change)
        
        # Contenedor derecho: Fila cabecera
        right_container = ttk.Frame(sheet_frame)
        right_container.pack(side=tk.RIGHT)
        
        ttk.Label(right_container, text="Fila cabecera:").pack(side=tk.LEFT)
        self.header_var = tk.StringVar(value="0")
        self.header_spin = ttk.Spinbox(right_container, from_=0, to=100, textvariable=self.header_var, width=5)
        self.header_spin.pack(side=tk.LEFT, padx=5)
        self.header_spin.bind("<ButtonRelease-1>", self._on_header_change)
        self.header_spin.bind("<KeyRelease>", self._on_header_change)
        
        # Mapeo de columnas
        columns_frame = ttk.LabelFrame(main_frame, text="Columnas", padding="5")
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.column_vars = {}
        self.column_combos = {}
        
        required_columns = ["Fecha Inicio", "Fecha Fin"]
        optional_columns = ["Fase", "Tareas", "Responsable"]
        
        ttk.Label(columns_frame, text="Obligatorias:", font=("", 9, "bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        row = 1
        for col in required_columns:
            ttk.Label(columns_frame, text=f"{col}:").grid(row=row, column=0, sticky=tk.W, pady=2)
            self.column_vars[col] = tk.StringVar()
            combo = ttk.Combobox(columns_frame, textvariable=self.column_vars[col], state="readonly", width=35)
            combo.grid(row=row, column=1, sticky=tk.W, pady=2, padx=5)
            self.column_combos[col] = combo
            row += 1
        
        ttk.Label(columns_frame, text="Opcionales:", font=("", 9, "bold")).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(10,0))
        row += 1
        
        for col in optional_columns:
            ttk.Label(columns_frame, text=f"{col}:").grid(row=row, column=0, sticky=tk.W, pady=2)
            self.column_vars[col] = tk.StringVar()
            combo = ttk.Combobox(columns_frame, textvariable=self.column_vars[col], state="readonly", width=35)
            combo.grid(row=row, column=1, sticky=tk.W, pady=2, padx=5)
            self.column_combos[col] = combo
            row += 1
        
        # Botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Cancelar", command=self._cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cargar", command=self._confirm).pack(side=tk.RIGHT)
        
    def _select_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.file_path = path
            self.file_label.config(text=path.split("/")[-1].split("\\")[-1])
            self._load_sheets()
    
    def _load_sheets(self):
        try:
            xl = pd.ExcelFile(self.file_path)
            self.sheets = xl.sheet_names
            self.sheet_combo["values"] = self.sheets
            if self.sheets:
                self.sheet_combo.current(0)
                self._on_sheet_change(None)
        except Exception as e:
            messagebox.showerror("Error", f"Error al leer el archivo: {e}")
    
    def _on_sheet_change(self, event):
        self._autodetect_header()
    
    def _autodetect_header(self):
        if not self.file_path or not self.sheet_var.get():
            return
        try:
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_var.get(), header=None, nrows=20)
            for idx, row in df.iterrows():
                non_empty = row.dropna()
                if len(non_empty) >= 2:
                    self.header_var.set(str(idx))
                    self._load_columns()
                    return
        except Exception:
            pass
    
    def _on_header_change(self, event=None):
        self.root.after(100, self._load_columns)
    
    def _load_columns(self):
        if not self.file_path or not self.sheet_var.get():
            return
        try:
            header_row = int(self.header_var.get())
            df = pd.read_excel(
                self.file_path, 
                sheet_name=self.sheet_var.get(), 
                header=header_row, 
                nrows=1
            )
            self.columns = ["(ninguno)"] + list(df.columns.astype(str))
            
            keywords = {
                "Fecha Inicio": ["inicio", "start", "fecha inicio"],
                "Fecha Fin": ["fin", "end", "fecha fin", "termino"],
                "Fase": ["fase", "phase", "etapa", "grupo"],
                "Tareas": ["tarea", "task", "actividad", "descripcion"],
                "Responsable": ["responsable", "owner", "asignado", "recurso"]
            }
            
            for col_name, combo in self.column_combos.items():
                combo["values"] = self.columns
                matched = False
                
                for excel_col in self.columns[1:]:
                    excel_lower = excel_col.lower()
                    for kw in keywords.get(col_name, []):
                        if kw in excel_lower:
                            combo.set(excel_col)
                            matched = True
                            break
                    if matched:
                        break
                
                if not matched:
                    combo.set("(ninguno)")
                        
        except Exception as e:
            print(f"Error cargando columnas: {e}")
    
    def _validate(self):
        if not self.file_path:
            messagebox.showwarning("Aviso", "Selecciona un archivo Excel")
            return False
        if not self.sheet_var.get():
            messagebox.showwarning("Aviso", "Selecciona una hoja")
            return False
        
        for required in ["Fecha Inicio", "Fecha Fin"]:
            val = self.column_vars[required].get()
            if not val or val == "(ninguno)":
                messagebox.showwarning("Aviso", f"La columna '{required}' es requerida")
                return False
        return True
    
    def _confirm(self):
        if not self._validate():
            return
        
        column_mapping = {}
        for col_name, var in self.column_vars.items():
            val = var.get()
            if val and val != "(ninguno)":
                column_mapping[col_name] = val
        
        config = {
            "file_path": self.file_path,
            "sheet_name": self.sheet_var.get(),
            "header": int(self.header_var.get()),
            "column_mapping": column_mapping
        }
        if self.on_load:
            self.on_load(config)
    
    def _cancel(self):
        self.root.destroy()
    
    def show(self):
        self.root.mainloop()

def show_excel_config(on_load=None):
    gui = ExcelConfigGUI(on_load=on_load)
    gui.show()