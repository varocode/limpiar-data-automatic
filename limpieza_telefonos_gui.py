import pandas as pd
import numpy as np
import re
import logging
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from collections import defaultdict
from datetime import datetime

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger('limpieza_telefonos')

class ColumnSelector(tk.Frame):
    def __init__(self, parent, title, selectmode=tk.SINGLE):
        tk.Frame.__init__(self, parent)
        self.title = title

        # Crear label para título
        self.label = tk.Label(self, text=title, font=("Arial", 10, "bold"))
        self.label.pack(pady=(10, 5))

        # Crear listbox con scrollbar
        self.frame = tk.Frame(self)
        self.scrollbar = tk.Scrollbar(self.frame)
        self.listbox = tk.Listbox(self.frame, height=10, width=30,
                                  selectmode=selectmode,
                                  yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.listbox.yview)

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Para el caso de MULTIPLE, agregar botones
        if selectmode == tk.MULTIPLE:
            self.btn_frame = tk.Frame(self)
            self.btn_up = tk.Button(self.btn_frame, text="↑", command=self.move_up)
            self.btn_down = tk.Button(self.btn_frame, text="↓", command=self.move_down)

            self.btn_up.pack(side=tk.LEFT, padx=5)
            self.btn_down.pack(side=tk.LEFT, padx=5)
            self.btn_frame.pack(pady=5)

    def get_selection(self):
        """Obtiene los elementos seleccionados"""
        if self.listbox.curselection():
            return [self.listbox.get(i) for i in self.listbox.curselection()]
        return []

    def get_all_items(self):
        """Obtiene todos los elementos en el listbox"""
        return [self.listbox.get(i) for i in range(self.listbox.size())]

    def populate(self, items):
        """Llena el listbox con los elementos dados"""
        self.listbox.delete(0, tk.END)
        for item in items:
            self.listbox.insert(tk.END, item)

    def move_up(self):
        """Mueve el elemento seleccionado hacia arriba"""
        if not self.listbox.curselection():
            return

        for i in self.listbox.curselection():
            if i > 0:
                text = self.listbox.get(i)
                self.listbox.delete(i)
                self.listbox.insert(i-1, text)
                self.listbox.select_set(i-1)
                self.listbox.see(i-1)

    def move_down(self):
        """Mueve el elemento seleccionado hacia abajo"""
        if not self.listbox.curselection():
            return

        sel = list(self.listbox.curselection())
        sel.reverse()  # Procesamos de abajo hacia arriba

        for i in sel:
            if i < self.listbox.size() - 1:
                text = self.listbox.get(i)
                self.listbox.delete(i)
                self.listbox.insert(i+1, text)
                self.listbox.select_set(i+1)
                self.listbox.see(i+1)

class WordReplacementDialog(tk.Toplevel):
    """Diálogo para reemplazar palabras en una columna"""
    def __init__(self, parent, df, columna):
        super().__init__(parent)
        self.title(f"Reemplazar palabras - {columna}")
        self.geometry("600x400")
        self.df = df
        self.columna = columna
        self.reemplazos = {}

        # Extraer valores únicos de la columna
        if columna in df.columns:
            valores = df[columna].dropna().astype(str).unique()
            valores = [v for v in valores if v.strip()]
        else:
            valores = []

        # Frame principal
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Label explicativo
        tk.Label(main_frame, text=f"Reemplazar palabras en la columna '{columna}'",
                font=("Arial", 10, "bold")).pack(pady=(0, 10))

        # Frame para la tabla de reemplazos
        table_frame = tk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Cabeceras
        headers_frame = tk.Frame(table_frame)
        headers_frame.pack(fill=tk.X, pady=5)

        tk.Label(headers_frame, text="Valor original", width=25, font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Label(headers_frame, text="Nuevo valor", width=25, font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)

        # Contenedor con scroll para los reemplazos
        self.canvas = tk.Canvas(table_frame)
        scrollbar = tk.Scrollbar(table_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Añadir campos para cada valor único
        self.entry_pairs = []
        for valor in valores:
            self.add_replacement_row(valor)

        # Botones de acción
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        tk.Button(button_frame, text="Aceptar", command=self.aceptar).pack(side=tk.RIGHT, padx=5)
        tk.Button(button_frame, text="Cancelar", command=self.cancelar).pack(side=tk.RIGHT, padx=5)

    def add_replacement_row(self, valor):
        """Añade una fila con entrada para reemplazo"""
        row_frame = tk.Frame(self.scrollable_frame)
        row_frame.pack(fill=tk.X, pady=2)

        # Valor original (solo lectura)
        lbl = tk.Label(row_frame, text=valor, width=25, anchor="w")
        lbl.pack(side=tk.LEFT, padx=5)

        # Nuevo valor (editable)
        entry = tk.Entry(row_frame, width=25)
        entry.insert(0, valor)  # Por defecto, el mismo valor
        entry.pack(side=tk.LEFT, padx=5)

        self.entry_pairs.append((valor, entry))

    def aceptar(self):
        """Guarda los reemplazos y cierra el diálogo"""
        for original, entry in self.entry_pairs:
            nuevo = entry.get().strip()
            if nuevo and nuevo != original:
                self.reemplazos[original] = nuevo

        self.destroy()

    def cancelar(self):
        """Cierra el diálogo sin guardar"""
        self.reemplazos = {}
        self.destroy()

    def get_replacements(self):
        """Retorna el diccionario de reemplazos"""
        return self.reemplazos

class ComentariosFrame(tk.LabelFrame):
    """Frame para generar columna de comentarios"""
    def __init__(self, parent, df):
        super().__init__(parent, text="Generador de Columna COMENTARIOS")
        self.df = df
        self.parent = parent
        self.reemplazos_columnas = {}  # Diccionario para almacenar reemplazos por columna

        self.create_widgets()

    def create_widgets(self):
        # Frame principal con scroll
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Crear canvas con scroll
        canvas = tk.Canvas(main_frame)
        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Sección de selección de columnas
        self.sections = []

        # 1. Tipo de producto
        self.add_section(scrollable_frame, "Tipo de producto", False)

        # 2. Límite pesos
        section_pesos = self.add_section(scrollable_frame, "Límite pesos", True, "RD$")

        # 3. Límite dólares
        section_dolares = self.add_section(scrollable_frame, "Límite dólares", True, "US$")

        # 4. Límite crédito diferido
        section_diferido = self.add_section(scrollable_frame, "Límite crédito diferido", True, "RD$")

        # Opción para reemplazar 0 por 1 en límite diferido
        zero_frame = tk.Frame(section_diferido)
        zero_frame.pack(fill=tk.X, pady=5)

        self.replace_zero_var = tk.BooleanVar(value=False)
        tk.Checkbutton(zero_frame, text="Reemplazar 0 por 1",
                       variable=self.replace_zero_var).pack(side=tk.LEFT, padx=5)

        # 5-11. Otros 1-7
        for i in range(1, 8):
            self.add_section(scrollable_frame, f"Otros {i}", False)

        # Sección de reemplazo de palabras
        replace_frame = tk.LabelFrame(scrollable_frame, text="Reemplazo de palabras")
        replace_frame.pack(fill=tk.X, pady=10, padx=5)

        replace_top = tk.Frame(replace_frame)
        replace_top.pack(fill=tk.X, pady=5)

        self.replace_enabled = tk.BooleanVar(value=False)
        tk.Checkbutton(replace_top, text="Habilitar reemplazo",
                     variable=self.replace_enabled,
                     command=self.update_order_list).pack(side=tk.LEFT, padx=5)

        # Columna para reemplazo
        tk.Label(replace_top, text="Columna:").pack(side=tk.LEFT, padx=5)
        self.replace_column = tk.StringVar()
        self.replace_combo = ttk.Combobox(replace_top, textvariable=self.replace_column, state="readonly")
        self.replace_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.replace_combo.bind("<<ComboboxSelected>>", lambda e: self.update_order_list())

        # Botón para abrir diálogo de reemplazo
        self.btn_reemplazar = tk.Button(replace_top, text="Configurar reemplazos",
                                      command=self.open_replacement_dialog)
        self.btn_reemplazar.pack(side=tk.LEFT, padx=5)

        # Sección para ordenar columnas
        order_frame = tk.LabelFrame(scrollable_frame, text="Orden de concatenación")
        order_frame.pack(fill=tk.X, pady=10, padx=5)

        # Lista de columnas seleccionadas
        list_frame = tk.Frame(order_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.order_listbox = tk.Listbox(list_frame, height=8)
        self.order_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # Botones para mover
        btn_frame = tk.Frame(list_frame)
        btn_frame.pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="↑", command=self.move_up).pack(fill=tk.X, pady=2)
        tk.Button(btn_frame, text="↓", command=self.move_down).pack(fill=tk.X, pady=2)

        # Botón para previsualizar
        preview_frame = tk.Frame(scrollable_frame)
        preview_frame.pack(fill=tk.X, pady=5)

        tk.Button(preview_frame, text="Previsualizar Comentarios",
                command=self.previsualizar_comentarios, bg="#2196F3", fg="white").pack(pady=5)

        # Separador
        ttk.Separator(scrollable_frame, orient='horizontal').pack(fill=tk.X, pady=10)

        # Frame para previsualización
        self.preview_frame = tk.LabelFrame(scrollable_frame, text="Previsualización de Comentarios")
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)

        # Área de texto para previsualización
        self.preview_text = tk.Text(self.preview_frame, height=5, width=60, wrap=tk.WORD)
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.preview_text.config(state=tk.DISABLED)

        # Botón para generar columna
        generate_frame = tk.Frame(scrollable_frame)
        generate_frame.pack(fill=tk.X, pady=10)

        tk.Button(generate_frame, text="Generar Columna COMENTARIOS",
                command=self.generar_comentarios, bg="#4CAF50", fg="white",
                font=("Arial", 10, "bold")).pack(pady=5)

        # Actualizar lista de columnas disponibles
        self.update_column_lists()

    def add_section(self, parent, title, is_money=False, prefix=""):
        """Añade una sección para seleccionar una columna"""
        section = tk.LabelFrame(parent, text=title)
        section.pack(fill=tk.X, pady=5, padx=5)

        # Si es un apartado "Otros", agregar campo de texto personalizado ARRIBA
        text_entry = None
        text_var = tk.StringVar()
        text_enabled = tk.BooleanVar(value=False)

        if "Otros" in title:
            # Frame para el campo de texto
            text_frame = tk.Frame(section)
            text_frame.pack(fill=tk.X, pady=5)

            # CheckBox para habilitar/deshabilitar texto personalizado
            # IMPORTANTE: ahora es independiente del estado de la sección
            text_check = tk.Checkbutton(text_frame, text="Texto personalizado:",
                         variable=text_enabled,
                         command=lambda: self.update_text_state(text_entry, text_enabled.get()))
            text_check.pack(side=tk.LEFT, padx=5)

            # Campo de texto
            text_entry = tk.Entry(text_frame, textvariable=text_var)
            text_entry["state"] = "normal" if text_enabled.get() else "disabled"
            text_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

            # Actualizar la lista de orden cuando cambia el texto
            text_var.trace_add("write", lambda *args: self.update_order_list())

        # Frame para selección de columna (ahora va DESPUÉS del texto en Otros)
        top_frame = tk.Frame(section)
        top_frame.pack(fill=tk.X, pady=5)

        # CheckBox para habilitar/deshabilitar
        enabled_var = tk.BooleanVar(value=False)
        check = tk.Checkbutton(top_frame, text="Habilitar",
                     variable=enabled_var,
                     command=lambda: self.update_section_state(combobox, enabled_var.get()))
        check.pack(side=tk.LEFT, padx=5)

        # Selector de columna
        tk.Label(top_frame, text="Columna:").pack(side=tk.LEFT, padx=5)
        column_var = tk.StringVar()
        combobox = ttk.Combobox(top_frame, textvariable=column_var, state="disabled")
        combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        combobox.bind("<<ComboboxSelected>>", lambda e: self.update_order_list())

        # Guardar configuración de sección
        section_config = {
            "frame": section,
            "enabled": enabled_var,
            "column": column_var,
            "combobox": combobox,
            "is_money": is_money,
            "prefix": prefix,
            "check": check,
            "text_entry": text_entry,
            "text_var": text_var,
            "text_enabled": text_enabled,
            "title": title
        }

        self.sections.append(section_config)
        return section

    def update_section_state(self, combobox, enabled):
        """Actualiza el estado del combobox según el checkbox"""
        combobox["state"] = "readonly" if enabled else "disabled"
        self.update_order_list()

    def update_text_state(self, text_entry, enabled):
        """Actualiza el estado del campo de texto según su checkbox"""
        if text_entry:
            text_entry["state"] = "normal" if enabled else "disabled"
            self.update_order_list()

    def update_order_list(self):
        """Actualiza la lista de orden según las secciones habilitadas"""
        self.order_listbox.delete(0, tk.END)

        for section in self.sections:
            # Primero agregamos el texto personalizado si está habilitado
            if section.get("text_entry") is not None and section["text_enabled"].get():
                texto = section["text_var"].get() or "[Texto vacío]"
                display_text = f"{section['title']}: \"{texto}\" (texto)"
                self.order_listbox.insert(tk.END, display_text)

            # Luego agregamos la columna si está habilitada
            if section["enabled"].get() and section["column"].get():
                display_text = f"{section['column'].get()}"
                if section["is_money"]:
                    display_text += f" ({section['prefix']})"
                self.order_listbox.insert(tk.END, display_text)

        # Añadir columna de reemplazo si está habilitada
        if self.replace_enabled.get() and self.replace_column.get():
            self.order_listbox.insert(tk.END, f"{self.replace_column.get()} (reemplazos)")

    def move_up(self):
        """Mueve el elemento seleccionado hacia arriba en la lista de orden"""
        if not self.order_listbox.curselection():
            return

        index = self.order_listbox.curselection()[0]
        if index > 0:
            text = self.order_listbox.get(index)
            self.order_listbox.delete(index)
            self.order_listbox.insert(index-1, text)
            self.order_listbox.selection_set(index-1)

    def move_down(self):
        """Mueve el elemento seleccionado hacia abajo en la lista de orden"""
        if not self.order_listbox.curselection():
            return

        index = self.order_listbox.curselection()[0]
        if index < self.order_listbox.size() - 1:
            text = self.order_listbox.get(index)
            self.order_listbox.delete(index)
            self.order_listbox.insert(index+1, text)
            self.order_listbox.selection_set(index+1)

    def open_replacement_dialog(self):
        """Abre el diálogo para configurar reemplazos"""
        columna = self.replace_column.get()
        if not columna:
            messagebox.showinfo("Información", "Primero selecciona una columna para reemplazar")
            return

        dialog = WordReplacementDialog(self, self.parent.df, columna)
        self.wait_window(dialog)

        replacements = dialog.get_replacements()
        if replacements:
            self.reemplazos_columnas[columna] = replacements
            messagebox.showinfo("Reemplazos guardados",
                              f"Se configuraron {len(replacements)} reemplazos para la columna '{columna}'")

            # Actualizar lista de orden
            self.update_order_list()

    def previsualizar_comentarios(self):
        """Muestra una previsualización de cómo quedarían los comentarios"""
        if not hasattr(self.parent, 'df') or self.parent.df is None:
            messagebox.showerror("Error", "No hay datos cargados")
            return

        # Verificar que hay al menos una columna seleccionada
        order_items = self.order_listbox.get(0, tk.END)
        if not order_items:
            messagebox.showinfo("Información", "No hay columnas seleccionadas para concatenar")
            return

        try:
            # Procesar columnas según orden para obtener una muestra
            comentarios_muestra = self.generar_comentarios_muestra(5)  # Muestra primeras 5 filas

            # Mostrar previsualización
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, comentarios_muestra)
            self.preview_text.config(state=tk.DISABLED)

        except Exception as e:
            messagebox.showerror("Error", f"Error al previsualizar comentarios: {str(e)}")
            logger.error(f"Error al previsualizar comentarios: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())

    def generar_comentarios_muestra(self, num_rows=5):
        """Genera una muestra de comentarios para previsualización"""
        # Crear copia del DataFrame para trabajar (solo las primeras filas)
        df_muestra = self.parent.df.head(num_rows).copy()

        # Obtener orden de columnas
        order_items = self.order_listbox.get(0, tk.END)

        # Procesar columnas según orden
        resultados = []

        for idx, row in df_muestra.iterrows():
            partes = []

            for item in order_items:
                # Manejar texto personalizado de apartados "Otros"
                if "(texto)" in item:
                    # Extraer el texto personalizado entre comillas
                    match = re.search(r'"([^"]*)"', item)
                    if match:
                        texto_personalizado = match.group(1)
                        if texto_personalizado and texto_personalizado != "[Texto vacío]":
                            partes.append(texto_personalizado)
                    continue

                # Extraer nombre de columna y tipo de procesamiento
                if " (" in item:
                    col_name, tipo = item.split(" (", 1)
                    tipo = tipo.rstrip(")")
                else:
                    col_name = item
                    tipo = "normal"

                # Verificar si la columna existe
                if col_name not in df_muestra.columns:
                    continue

                # Obtener valor
                valor = row[col_name]

                # Procesar según tipo
                if pd.isna(valor):
                    valor_proc = ""
                else:
                    valor_proc = str(valor)

                    if tipo == "RD$":
                        valor_proc = f"RD${valor_proc}"

                    elif tipo == "US$":
                        valor_proc = f"US${valor_proc}"

                    elif tipo == "reemplazos":
                        # Aplicar reemplazos si existen
                        if col_name in self.reemplazos_columnas:
                            for original, nuevo in self.reemplazos_columnas[col_name].items():
                                valor_proc = valor_proc.replace(original, nuevo)

                # Caso especial: reemplazar 0 por 1 en límite diferido
                for section in self.sections:
                    if (section["enabled"].get() and section["column"].get() == col_name and
                        "Límite crédito diferido" in section["frame"]["text"] and
                        self.replace_zero_var.get()):
                        # Reemplazar valores "0" o "RD$0" por "1" o "RD$1"
                        if tipo == "RD$" and valor_proc == "RD$0":
                            valor_proc = "RD$1"
                        elif valor_proc == "0":
                            valor_proc = "1"

                # Añadir a partes si no está vacío
                if valor_proc:
                    partes.append(valor_proc)

            # Concatenar partes con un espacio
            comentario = " ".join(partes)
            resultados.append(f"Fila {idx+1}: {comentario}")

        return "\n\n".join(resultados)

    def generar_comentarios(self):
        """Genera la columna COMENTARIOS según la configuración"""
        if not hasattr(self.parent, 'df') or self.parent.df is None:
            messagebox.showerror("Error", "No hay datos cargados")
            return

        # Verificar que hay al menos una columna seleccionada
        order_items = self.order_listbox.get(0, tk.END)
        if not order_items:
            messagebox.showinfo("Información", "No hay columnas seleccionadas para concatenar")
            return

        # Pedir confirmación mostrando previsualización
        self.previsualizar_comentarios()
        if not messagebox.askyesno("Confirmar", "¿Desea generar la columna COMENTARIOS con esta configuración?"):
            return

        try:
            logger.info("=== INICIANDO GENERACIÓN DE COLUMNA COMENTARIOS ===")

            # Crear copia del DataFrame para trabajar
            logger.info("Creando copia del DataFrame original para añadir comentarios")
            df_trabajo = self.parent.df.copy()

            # Inicializar columna COMENTARIOS
            df_trabajo["COMENTARIOS"] = ""
            logger.info("Columna 'COMENTARIOS' inicializada en el DataFrame")

            # Procesar filas una por una
            total_filas = len(df_trabajo)
            logger.info(f"Procesando {total_filas} filas para generar comentarios...")

            for idx, row in df_trabajo.iterrows():
                if idx % 1000 == 0 and idx > 0:  # Log cada 1000 filas
                    logger.info(f"Procesadas {idx} de {total_filas} filas...")

                partes = []

                for item in order_items:
                    # Manejar texto personalizado de apartados "Otros"
                    if "(texto)" in item:
                        # Extraer el texto personalizado entre comillas
                        match = re.search(r'"([^"]*)"', item)
                        if match:
                            texto_personalizado = match.group(1)
                            if texto_personalizado and texto_personalizado != "[Texto vacío]":
                                partes.append(texto_personalizado)
                        continue

                    # Extraer nombre de columna y tipo de procesamiento
                    if " (" in item:
                        col_name, tipo = item.split(" (", 1)
                        tipo = tipo.rstrip(")")
                    else:
                        col_name = item
                        tipo = "normal"

                    # Verificar si la columna existe
                    if col_name not in df_trabajo.columns:
                        continue

                    # Obtener valor
                    valor = row[col_name]

                    # Procesar según tipo
                    if pd.isna(valor):
                        valor_proc = ""
                    else:
                        valor_proc = str(valor)

                        if tipo == "RD$":
                            valor_proc = f"RD${valor_proc}"

                        elif tipo == "US$":
                            valor_proc = f"US${valor_proc}"

                        elif tipo == "reemplazos":
                            # Aplicar reemplazos si existen
                            if col_name in self.reemplazos_columnas:
                                for original, nuevo in self.reemplazos_columnas[col_name].items():
                                    valor_proc = valor_proc.replace(original, nuevo)

                    # Caso especial: reemplazar 0 por 1 en límite diferido
                    for section in self.sections:
                        if (section["enabled"].get() and section["column"].get() == col_name and
                            "Límite crédito diferido" in section["frame"]["text"] and
                            self.replace_zero_var.get()):
                            # Reemplazar valores "0" o "RD$0" por "1" o "RD$1"
                            if tipo == "RD$" and valor_proc == "RD$0":
                                valor_proc = "RD$1"
                            elif valor_proc == "0":
                                valor_proc = "1"

                    # Añadir a partes si no está vacío
                    if valor_proc:
                        partes.append(valor_proc)

                # Concatenar partes con un espacio
                df_trabajo.at[idx, "COMENTARIOS"] = " ".join(partes)

            # Guardar DataFrame con columna COMENTARIOS
            logger.info("Columna COMENTARIOS generada exitosamente")

            # Ejemplos de comentarios generados
            ejemplos = df_trabajo["COMENTARIOS"].head(3).tolist()
            logger.info(f"Ejemplos de comentarios generados:")
            for i, ejemplo in enumerate(ejemplos):
                logger.info(f"  Fila {i+1}: {ejemplo}")

            # Guardar en la instancia principal
            self.parent.df_comentarios = df_trabajo
            self.parent.comentarios_generados = True
            logger.info("DataFrame con comentarios guardado en memoria y listo para fusión")

            # Mostrar mensaje de éxito
            messagebox.showinfo("Éxito", "Columna COMENTARIOS generada correctamente y lista para fusión")

        except Exception as e:
            messagebox.showerror("Error", f"Error al generar comentarios: {str(e)}")
            logger.error(f"Error al generar comentarios: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())

    def fusionar_con_base_original(self, df_limpio, cedula_col, tel_cols_seleccionadas):
        """Fusiona la base limpia con una copia de la base original completa"""
        try:
            logger.info("\n=== INICIANDO PROCESO DE FUSIÓN ===")

            # Verificar que tenemos los datos necesarios
            if df_limpio is None:
                messagebox.showerror("Error", "No hay datos limpios para fusionar")
                return

            if cedula_col is None or not tel_cols_seleccionadas:
                messagebox.showerror("Error", "Falta información de columnas para fusionar")
                return

            # 1. Crear COPIA EXPLÍCITA de la base original
            logger.info("Creando una copia de la base original para la fusión")

            # Determinar qué base usar (original o con comentarios)
            if hasattr(self, 'comentarios_generados') and self.comentarios_generados and hasattr(self, 'df_comentarios'):
                logger.info("✓ COMENTARIOS DETECTADOS: Se usará la base con la columna COMENTARIOS")
                df_base_completa = self.df_comentarios.copy()
                # Verificar que la columna COMENTARIOS existe
                if "COMENTARIOS" in df_base_completa.columns:
                    logger.info(f"✓ Columna COMENTARIOS encontrada con {len(df_base_completa)} filas")
                    # Mostrar ejemplos de comentarios
                    ejemplos = df_base_completa["COMENTARIOS"].head(3).tolist()
                    logger.info(f"Ejemplos de comentarios que serán incluidos en la fusión:")
                    for i, ejemplo in enumerate(ejemplos):
                        logger.info(f"  Fila {i+1}: {ejemplo}")
                else:
                    logger.warning("⚠ La columna COMENTARIOS no se encontró en el DataFrame")
            else:
                logger.info("No se detectaron comentarios generados: Se usará la base original")
                df_base_completa = self.df.copy()

            # 2. Eliminar las columnas de teléfono de la COPIA
            logger.info(f"Eliminando las {len(tel_cols_seleccionadas)} columnas seleccionadas de la COPIA...")
            df_base_sin_tel = df_base_completa.drop(columns=tel_cols_seleccionadas)

            # 3. Renombrar columnas de base limpia para evitar conflictos
            # Si tenemos columnas con el mismo nombre, renombramos las de la base limpia
            suffix = "_LIMPIO"
            df_limpio_renamed = df_limpio.copy()

            # Si la columna de cédula ya existe, la renombramos
            if cedula_col in df_base_sin_tel.columns:
                df_limpio_renamed = df_limpio_renamed.rename(columns={cedula_col: f"{cedula_col}{suffix}"})

            # 4. Solicitar al usuario donde guardar la fusión
            file_types = (("Excel files", "*.xlsx"), ("All files", "*.*"))
            output_fusion = filedialog.asksaveasfilename(
                title="Guardar base fusionada como",
                defaultextension=".xlsx",
                filetypes=file_types,
                initialfile=f"BaseFusionada_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            )

            if not output_fusion:
                logger.info("Operación de fusión cancelada por el usuario")
                return

            # 5. Crear la fusión: dos opciones posibles
            logger.info("Fusionando COPIA con base limpia...")

            # Preguntar cómo quiere hacer la fusión
            fusion_tipo = messagebox.askquestion("Tipo de fusión",
                "¿Deseas fusionar las bases por cédula (SÍ) o simplemente añadir columnas (NO)?")

            if fusion_tipo == 'yes':  # Fusión por cédula
                # Buscar cómo se llama la columna de cédula en df_limpio_renamed
                cedula_limpia = cedula_col
                if f"{cedula_col}{suffix}" in df_limpio_renamed.columns:
                    cedula_limpia = f"{cedula_col}{suffix}"

                # Fusionar bases por columna de cédula
                df_fusionado = pd.merge(
                    df_base_sin_tel,
                    df_limpio_renamed,
                    left_on=cedula_col,
                    right_on=cedula_limpia,
                    how='left'
                )

                # Eliminar columna duplicada de cédula si existe
                if cedula_limpia != cedula_col and cedula_limpia in df_fusionado.columns:
                    df_fusionado = df_fusionado.drop(columns=[cedula_limpia])

                logger.info(f"Fusión realizada por coincidencia de cédulas")
            else:  # Añadir columnas sin fusionar
                # Verificar que tenemos el mismo número de filas
                if len(df_base_sin_tel) != len(df_limpio_renamed):
                    logger.warning(f"¡Advertencia! Las bases tienen diferente cantidad de filas: "
                                  f"Copia de base: {len(df_base_sin_tel)}, "
                                  f"Base limpia: {len(df_limpio_renamed)}")

                # Simplemente concatenar columnas
                df_fusionado = pd.concat([df_base_sin_tel, df_limpio_renamed], axis=1)
                logger.info(f"Columnas añadidas a la copia de la base")

            # 6. Guardar resultado
            logger.info(f"Guardando base fusionada como: {output_fusion}")

            # Verificar si la columna COMENTARIOS está en el resultado final
            if "COMENTARIOS" in df_fusionado.columns:
                logger.info("✓ La columna COMENTARIOS se ha incluido correctamente en la fusión")
            else:
                logger.warning("⚠ La columna COMENTARIOS no está presente en la fusión final")

            df_fusionado.to_excel(output_fusion, index=False)
            logger.info("✅ Fusión completada exitosamente")

            # Mostrar mensaje de éxito
            messagebox.showinfo("Éxito", "Fusión completada exitosamente")

        except Exception as e:
            logger.error(f"Error durante la fusión: {str(e)}")
            messagebox.showerror("Error", f"Error durante la fusión: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())

    def update_column_lists(self):
        """Actualiza las listas de columnas disponibles"""
        if hasattr(self.parent, 'df') and self.parent.df is not None:
            columns = self.parent.df.columns.tolist()

            # Actualizar cada sección
            for section in self.sections:
                section["combobox"]["values"] = columns

            # Actualizar combobox de reemplazo
            self.replace_combo["values"] = columns

            logger.info(f"Listas de columnas actualizadas en el generador de comentarios: {len(columns)} columnas")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Limpieza de Teléfonos y Generador de Comentarios")
        self.geometry("900x700")
        self.df = None
        self.df_comentarios = None
        self.df_limpio = None  # Para almacenar el resultado de la limpieza
        self.columnas_limpias = []  # Para almacenar las columnas que se limpiaron
        self.cedula_col = None  # Para almacenar la columna de cédula
        self.comentarios_generados = False

        self.create_widgets()

    def create_widgets(self):
        # Crear notebook (pestañas)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Pestaña 1: Limpieza de teléfonos
        self.tab_limpieza = tk.Frame(self.notebook)
        self.notebook.add(self.tab_limpieza, text="Limpieza de Teléfonos")

        # Pestaña 2: Generador de comentarios
        self.tab_comentarios = tk.Frame(self.notebook)
        self.notebook.add(self.tab_comentarios, text="Generador de Comentarios")

        # Contenido de pestaña 1 (Limpieza)
        self.create_limpieza_widgets()

        # Contenido de pestaña 2 (Comentarios) - Se creará después de cargar datos

        # Área de log (común para ambas pestañas)
        log_frame = tk.LabelFrame(self, text="Log")
        log_frame.pack(fill=tk.X, expand=False, padx=10, pady=10)

        self.log_text = tk.Text(log_frame, height=10, width=80)
        self.log_text.pack(fill=tk.X, expand=True, padx=5, pady=5)

        # Redireccionar los logs a este widget
        self.text_handler = TextHandler(self.log_text)
        logger.addHandler(self.text_handler)

    def create_limpieza_widgets(self):
        # Frame principal
        main_frame = tk.Frame(self.tab_limpieza)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Frame superior para selección de archivo
        file_frame = tk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)

        tk.Label(file_frame, text="Archivo Excel:").pack(side=tk.LEFT, padx=5)
        self.file_entry = tk.Entry(file_frame, width=50)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        tk.Button(file_frame, text="Buscar", command=self.browse_file).pack(side=tk.LEFT, padx=5)
        tk.Button(file_frame, text="Cargar", command=self.load_file).pack(side=tk.LEFT, padx=5)

        # Frame para selectores de columnas
        columns_frame = tk.Frame(main_frame)
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Selector de columnas disponibles
        self.available_cols = ColumnSelector(columns_frame, "Columnas disponibles", tk.MULTIPLE)
        self.available_cols.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # Frame para botones de transferencia
        btn_frame = tk.Frame(columns_frame)
        btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5)

        tk.Button(btn_frame, text="-->", command=self.move_to_selected).pack(pady=10)
        tk.Button(btn_frame, text="<--", command=self.move_to_available).pack(pady=10)

        # Selector de columna de cédula
        self.cedula_selector = ColumnSelector(columns_frame, "Columna de cédula")
        self.cedula_selector.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # Selector de columnas de teléfono
        self.tel_selector = ColumnSelector(columns_frame, "Columnas de teléfono (ordenadas)", tk.MULTIPLE)
        self.tel_selector.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # Frame para botones de acción
        action_frame = tk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=10)

        tk.Button(action_frame, text="Ejecutar limpieza", command=self.ejecutar_limpieza,
                  bg="#4CAF50", fg="white", font=("Arial", 10, "bold")).pack(side=tk.RIGHT, padx=5)

    def browse_file(self):
        """Abre diálogo para seleccionar archivo Excel"""
        filetypes = (("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        filename = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=filetypes)
        if filename:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filename)

    def load_file(self):
        """Carga el archivo Excel y muestra columnas"""
        filename = self.file_entry.get().strip()
        if not filename:
            messagebox.showerror("Error", "Por favor seleccione un archivo")
            return

        try:
            logger.info(f"Cargando archivo: {filename}")
            self.df = pd.read_excel(filename)

            # Actualizar lista de columnas disponibles en la pestaña de limpieza
            self.available_cols.populate(self.df.columns.tolist())
            self.cedula_selector.populate([])
            self.tel_selector.populate([])

            # Crear/actualizar el frame de comentarios
            self.create_comentarios_widgets()

            # Actualizar las listas de columnas en el frame de comentarios si ya existe
            if hasattr(self, 'comentarios_frame') and self.comentarios_frame is not None:
                try:
                    self.comentarios_frame.update_column_lists()
                except Exception as e:
                    logger.error(f"Error al actualizar columnas en comentarios: {str(e)}")

            logger.info(f"Archivo cargado correctamente. Columnas: {len(self.df.columns)}")
        except Exception as e:
            logger.error(f"Error al cargar archivo: {str(e)}")
            messagebox.showerror("Error", f"Error al cargar archivo: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())

    def create_comentarios_widgets(self):
        """Crea widgets para la pestaña de comentarios"""
        # Limpiar pestaña si ya tenía widgets
        for widget in self.tab_comentarios.winfo_children():
            widget.destroy()

        # Crear frame de comentarios
        self.comentarios_frame = ComentariosFrame(self, self.df)
        self.comentarios_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def move_to_selected(self):
        """Mueve las columnas seleccionadas a los selectores correspondientes"""
        selected = self.available_cols.get_selection()
        if not selected:
            return

        # Si no hay columna de cédula seleccionada, la primera selección va ahí
        if self.cedula_selector.listbox.size() == 0 and len(selected) > 0:
            self.cedula_selector.populate([selected[0]])
            selected = selected[1:]  # Removemos el primero que ya asignamos a cédula

        # El resto van a columnas de teléfono
        current_tel = self.tel_selector.get_all_items()
        self.tel_selector.populate(current_tel + selected)

        # Actualizamos disponibles
        current_avail = self.available_cols.get_all_items()
        self.available_cols.populate([c for c in current_avail if c not in selected])

    def move_to_available(self):
        """Mueve las columnas seleccionadas de vuelta a disponibles"""
        # Verificamos selecciones en ambos selectores
        ced_selected = self.cedula_selector.get_selection()
        tel_selected = self.tel_selector.get_selection()

        if not ced_selected and not tel_selected:
            return

        # Mover de cédula a disponibles
        if ced_selected:
            current_avail = self.available_cols.get_all_items()
            self.available_cols.populate(current_avail + ced_selected)
            self.cedula_selector.populate([])

        # Mover de teléfono a disponibles
        if tel_selected:
            current_avail = self.available_cols.get_all_items()
            self.available_cols.populate(current_avail + tel_selected)
            current_tel = self.tel_selector.get_all_items()
            self.tel_selector.populate([c for c in current_tel if c not in tel_selected])

    def ejecutar_limpieza(self):
        """Ejecuta el proceso de limpieza con las columnas seleccionadas"""
        if self.df is None:
            messagebox.showerror("Error", "Primero debe cargar un archivo")
            return

        # Verificar que tenemos columna de cédula y al menos una de teléfono
        cedula_cols = self.cedula_selector.get_all_items()
        tel_cols = self.tel_selector.get_all_items()

        if not cedula_cols:
            messagebox.showerror("Error", "Debe seleccionar una columna de cédula")
            return

        if not tel_cols:
            messagebox.showerror("Error", "Debe seleccionar al menos una columna de teléfono")
            return

        cedula_col = cedula_cols[0]

        # Guardar información para uso posterior
        self.cedula_col = cedula_col
        self.columnas_limpias = tel_cols

        # Preguntar dónde guardar el resultado de la limpieza
        file_types = (("Excel files", "*.xlsx"), ("All files", "*.*"))
        output_file = filedialog.asksaveasfilename(
            title="Guardar archivo limpio como",
            defaultextension=".xlsx",
            filetypes=file_types,
            initialfile=f"BaseLimpia_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        if not output_file:
            return  # Usuario canceló

        try:
            # Ejecutar limpieza
            self.df_limpio = self.procesar_limpieza(cedula_col, tel_cols, output_file)
            messagebox.showinfo("Éxito", "Proceso de limpieza completado exitosamente")

            # Preguntar si quiere fusionar con la base original
            if messagebox.askyesno("Fusionar bases",
                                "¿Deseas unir la base limpia con una copia de la base original\n"
                                "y reemplazar los teléfonos antiguos?"):
                # Aquí usamos una función a nivel de módulo en lugar de un método
                realizar_fusion(self)

        except Exception as e:
            logger.error(f"Error durante la limpieza: {str(e)}")
            messagebox.showerror("Error", f"Error durante la limpieza: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())

    def procesar_limpieza(self, cedula_col, tel_cols, output_file):
        """Ejecuta el proceso de limpieza con las columnas seleccionadas"""
        logger.info("Iniciando proceso de limpieza...")

        # Crear copia del DataFrame original con solo las columnas seleccionadas
        df_trabajo = self.df[[cedula_col] + tel_cols].copy()

        # ===========================
        # 1. NORMALIZAR TELÉFONOS (ELIMINAR 1 INICIAL)
        # ===========================
        logger.info("Normalizando números de teléfono (eliminando 1 inicial)...")

        telefonos_normalizados = 0

        def normalizar_telefono(valor):
            if not es_valor_valido(valor):
                return np.nan

            # Convertir a string y eliminar espacios
            telefono = str(valor).strip()

            # Eliminar todos los caracteres no numéricos
            solo_numeros = re.sub(r'\D', '', telefono)

            # Si comienza con 1 y tiene más de 1 dígito, eliminar el 1 inicial
            if solo_numeros.startswith('1') and len(solo_numeros) > 1:
                nonlocal telefonos_normalizados
                telefonos_normalizados += 1
                return solo_numeros[1:]

            return solo_numeros

        # Aplicar normalización a todas las columnas de teléfono
        for col in tel_cols:
            df_trabajo[col] = df_trabajo[col].apply(normalizar_telefono)

        logger.info(f"Se normalizaron {telefonos_normalizados} números de teléfono (se eliminó el 1 inicial)")

        # ===========================
        # 2. ELIMINAR TELÉFONOS DUPLICADOS GLOBALES
        # ===========================
        logger.info("Eliminando valores duplicados globalmente...")

        # Recolectar todos los valores en un solo Series
        valores_tel = []
        for col in tel_cols:
            for valor in df_trabajo[col]:
                if es_valor_valido(valor):
                    valores_tel.append(str(valor).strip())

        # Crear series y obtener valores únicos
        valores_tel_series = pd.Series(valores_tel)
        total_antes = len(valores_tel)
        unicos = valores_tel_series.drop_duplicates().tolist()
        total_unicos = len(unicos)

        logger.info(f"Valores totales: {total_antes}")
        logger.info(f"Valores únicos: {total_unicos}")
        logger.info(f"Duplicados eliminados: {total_antes - total_unicos}")

        # Conjunto para búsquedas rápidas
        set_unicos = set(unicos)
        usados = set()

        def limpiar_duplicados(celda):
            if not es_valor_valido(celda):
                return np.nan

            val = str(celda).strip()
            if val in usados:
                return np.nan
            if val in set_unicos:
                usados.add(val)
                return val
            return np.nan

        # Aplicar limpieza a todas las columnas de teléfono
        for col in tel_cols:
            df_trabajo[col] = df_trabajo[col].apply(limpiar_duplicados)

        # ===========================
        # 3. VALIDAR NÚMEROS DE TELÉFONO (DOMINICANOS)
        # ===========================
        logger.info("Validando números de teléfono dominicanos...")

        prefijos_validos = ['809', '829', '849']
        telefonos_invalidos = 0

        def validar_telefono(valor):
            if pd.isna(valor):
                return np.nan

            # Convertir a string y eliminar espacios
            telefono = str(valor).strip()

            # Validar longitud y prefijo
            if len(telefono) == 10 and telefono[:3] in prefijos_validos:
                return telefono
            else:
                nonlocal telefonos_invalidos
                telefonos_invalidos += 1
                return np.nan

        # Aplicar validación a todas las columnas de teléfono
        for col in tel_cols:
            df_trabajo[col] = df_trabajo[col].apply(validar_telefono)

        logger.info(f"Se encontraron y eliminaron {telefonos_invalidos} números de teléfono inválidos")

        # ===========================
        # 4. RELLENO DE CELDAS VACÍAS CON VALOR MÁS PRÓXIMO A LA DERECHA
        # ===========================
        logger.info("Rellenando celdas vacías con valores vecinos a la derecha...")

        # Definir función de relleno
        def rellenar_fila(fila):
            # Convertimos a lista para manipulación
            fila = list(fila)

            # Variable para contar celdas rellenadas en esta fila
            celdas_rellenadas_fila = 0

            # Saltamos la primera columna (cédula)
            for i in range(1, len(fila)):
                # Si la celda está vacía
                if pd.isna(fila[i]) or str(fila[i]).strip() == "":
                    # Buscamos el próximo valor no vacío a la derecha
                    for j in range(i+1, len(fila)):
                        if es_valor_valido(fila[j]):
                            # Movemos el valor
                            fila[i] = fila[j]
                            fila[j] = np.nan
                            celdas_rellenadas_fila += 1
                            break

            return fila, celdas_rellenadas_fila

        # Inicializar contador total
        total_celdas_rellenadas = 0

        # Aplicar función de relleno a cada fila y mantener conteo
        resultados = df_trabajo.apply(rellenar_fila, axis=1)
        df_rellenado = pd.DataFrame([r[0] for r in resultados], columns=df_trabajo.columns)
        total_celdas_rellenadas = sum(r[1] for r in resultados)

        # Reemplazar el DataFrame de trabajo
        df_trabajo = df_rellenado
        logger.info(f"Se rellenaron {total_celdas_rellenadas} celdas vacías")

        # ===========================
        # 5. GUARDAR RESULTADO FINAL
        # ===========================
        logger.info(f"Guardando archivo limpio como: {output_file}")
        df_trabajo.to_excel(output_file, index=False)
        logger.info("✅ Proceso completado exitosamente")

        return df_trabajo

# Funciones auxiliares
def es_valor_valido(valor):
    """Verifica si un valor no está vacío"""
    if pd.isna(valor):
        return False

    valor_str = str(valor).strip()
    return valor_str != ""

# Manejador para redirigir los logs al widget Text
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        logging.Handler.__init__(self)
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')
        self.text_widget.update()

# Función a nivel de módulo para realizar la fusión
def realizar_fusion(app):
    """Realiza la fusión entre la base limpia y una copia de la base original"""
    try:
        logger.info("\n=== INICIANDO PROCESO DE FUSIÓN ===")

        # Verificar que tenemos los datos necesarios
        if app.df_limpio is None:
            messagebox.showerror("Error", "No hay datos limpios para fusionar")
            return

        if app.cedula_col is None or not app.columnas_limpias:
            messagebox.showerror("Error", "Falta información de columnas para fusionar")
            return

        # 1. Determinar qué base usar como origen
        if app.comentarios_generados and app.df_comentarios is not None:
            logger.info("✓ COMENTARIOS DETECTADOS: Se usará la base con la columna COMENTARIOS")
            df_origen = app.df_comentarios.copy()
            # Verificar que la columna COMENTARIOS existe
            if "COMENTARIOS" in df_origen.columns:
                logger.info(f"✓ Columna COMENTARIOS encontrada con {len(df_origen)} filas")
                # Mostrar ejemplos de comentarios
                ejemplos = df_origen["COMENTARIOS"].head(3).tolist()
                logger.info(f"Ejemplos de comentarios que serán incluidos en la fusión:")
                for i, ejemplo in enumerate(ejemplos):
                    logger.info(f"  Fila {i+1}: {ejemplo}")
            else:
                logger.warning("⚠ La columna COMENTARIOS no se encontró en el DataFrame")
        else:
            logger.info("No se detectaron comentarios generados: Se usará la base original")
            df_origen = app.df.copy()

        # 2. Eliminar columnas de teléfono de la copia
        logger.info(f"Eliminando columnas de teléfono de la copia...")
        columnas_a_eliminar = [col for col in app.columnas_limpias if col in df_origen.columns]
        df_sin_telefonos = df_origen.drop(columns=columnas_a_eliminar)

        # 3. Preparar el DataFrame limpio
        df_limpio = app.df_limpio.copy()

        # 4. Solicitar al usuario donde guardar
        file_types = (("Excel files", "*.xlsx"), ("All files", "*.*"))
        output_fusion = filedialog.asksaveasfilename(
            title="Guardar base fusionada como",
            defaultextension=".xlsx",
            filetypes=file_types,
            initialfile=f"BaseFusionada_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        if not output_fusion:
            logger.info("Operación de fusión cancelada por el usuario")
            return

        # 5. Crear la fusión
        logger.info("Fusionando bases de datos...")

        # Preguntar cómo quiere hacer la fusión
        fusion_tipo = messagebox.askquestion("Tipo de fusión",
            "¿Deseas fusionar las bases por cédula (SÍ) o simplemente añadir columnas (NO)?")

        if fusion_tipo == 'yes':  # Fusión por cédula
            # Fusionar bases por columna de cédula
            df_fusionado = pd.merge(
                df_sin_telefonos,
                df_limpio,
                on=app.cedula_col,  # Usamos la misma columna en ambos DataFrames
                how='left'
            )

            logger.info(f"Fusión realizada por coincidencia de cédulas")
        else:  # Añadir columnas sin fusionar
            # Verificar que tenemos el mismo número de filas
            if len(df_sin_telefonos) != len(df_limpio):
                logger.warning(f"¡Advertencia! Las bases tienen diferente cantidad de filas: "
                              f"Base origen: {len(df_sin_telefonos)}, "
                              f"Base limpia: {len(df_limpio)}")

                # Preguntar si quiere continuar
                if not messagebox.askyesno("Advertencia",
                                 f"Las bases tienen diferente cantidad de filas:\n"
                                 f"Base origen: {len(df_sin_telefonos)}, "
                                 f"Base limpia: {len(df_limpio)}\n\n"
                                 f"¿Deseas continuar de todos modos?"):
                    return

            # Simplemente concatenar columnas
            df_fusionado = pd.concat([df_sin_telefonos, df_limpio.drop(columns=[app.cedula_col])], axis=1)
            logger.info(f"Columnas añadidas a la base de origen")

        # 6. Guardar resultado
        logger.info(f"Guardando base fusionada como: {output_fusion}")

        # Verificar si la columna COMENTARIOS está en el resultado final
        if "COMENTARIOS" in df_fusionado.columns:
            logger.info("✓ La columna COMENTARIOS se ha incluido correctamente en la fusión")
        else:
            logger.warning("⚠ La columna COMENTARIOS no está presente en la fusión final")

        df_fusionado.to_excel(output_fusion, index=False)
        logger.info("✅ Fusión completada exitosamente")

        # Mostrar mensaje de éxito
        messagebox.showinfo("Éxito", "Fusión completada exitosamente")

    except Exception as e:
        logger.error(f"Error durante la fusión: {str(e)}")
        messagebox.showerror("Error", f"Error durante la fusión: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())

# Punto de entrada principal
if __name__ == "__main__":
    app = App()
    app.mainloop()
