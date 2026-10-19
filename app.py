"""
Aplicación Gráfica (GUI) para la asignación y pegado de calificaciones en SIIUPP desde Excel.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core import BrowserManager, ExcelManager, clean_matricula


class SIIUPPApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Asignador de Calificaciones SIIUPP")
        self.root.geometry("980x750")
        self.root.minsize(850, 600)

        # Managers
        self.browser_manager = BrowserManager()
        self.upp_students: list[dict] = []
        self.excel_file_path: str = ""
        self.df_excel = None
        self.matched_data: list[dict] = []

        # Estilo visual
        self._configure_styles()

        # Construir Interfaz
        self._build_ui()

    def _configure_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Colores
        self.COLOR_BG = "#F4F6F9"
        self.COLOR_CARD = "#FFFFFF"
        self.COLOR_PRIMARY = "#1E40AF"      # Azul profesional
        self.COLOR_PRIMARY_HOVER = "#1D4ED8"
        self.COLOR_SUCCESS = "#059669"      # Verde
        self.COLOR_TEXT = "#1F2937"
        self.COLOR_MUTED = "#6B7280"
        self.COLOR_BORDER = "#E5E7EB"

        self.root.configure(bg=self.COLOR_BG)

        self.style.configure(".", background=self.COLOR_BG, foreground=self.COLOR_TEXT, font=("Segoe UI", 9))
        self.style.configure("Card.TFrame", background=self.COLOR_CARD, relief="flat")
        self.style.configure("Header.TLabel", background=self.COLOR_PRIMARY, foreground="#FFFFFF", font=("Segoe UI", 14, "bold"))
        self.style.configure("Subheader.TLabel", background=self.COLOR_PRIMARY, foreground="#E0E7FF", font=("Segoe UI", 9))
        self.style.configure("SectionTitle.TLabel", background=self.COLOR_CARD, font=("Segoe UI", 10, "bold"), foreground=self.COLOR_PRIMARY)
        self.style.configure("Status.TLabel", font=("Segoe UI", 9))

        # Botones
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 9, "bold"),
            background=self.COLOR_PRIMARY,
            foreground="#FFFFFF",
            padding=6
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", self.COLOR_PRIMARY_HOVER)],
            foreground=[("active", "#FFFFFF")]
        )

        self.style.configure(
            "Success.TButton",
            font=("Segoe UI", 10, "bold"),
            background=self.COLOR_SUCCESS,
            foreground="#FFFFFF",
            padding=8
        )
        self.style.map(
            "Success.TButton",
            background=[("active", "#047857")],
            foreground=[("active", "#FFFFFF")]
        )

        # Treeview
        self.style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", rowheight=24)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#E5E7EB", foreground=self.COLOR_TEXT)

    def _build_ui(self):
        # 1. Encabezado Superior
        header_frame = tk.Frame(self.root, bg=self.COLOR_PRIMARY, height=65)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False)

        lbl_title = ttk.Label(header_frame, text="🎓 Asignador de Calificaciones SIIUPP", style="Header.TLabel")
        lbl_title.pack(anchor="w", padx=20, pady=(8, 0))

        lbl_sub = ttk.Label(header_frame, text="Automatización segura de captura de notas desde Excel", style="Subheader.TLabel")
        lbl_sub.pack(anchor="w", padx=20, pady=(0, 8))

        # Contenedor principal con scroll o paneles
        main_container = tk.Frame(self.root, bg=self.COLOR_BG)
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        # Panel Superior: Paso 1 y Paso 2 (Horizontal o 2 Cards)
        top_panels = tk.Frame(main_container, bg=self.COLOR_BG)
        top_panels.pack(fill=tk.X, pady=(0, 10))

        # Card 1: Control de Navegador
        card1 = ttk.Frame(top_panels, style="Card.TFrame", padding=12)
        card1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        ttk.Label(card1, text="1. Navegador Chrome (SIIUPP)", style="SectionTitle.TLabel").pack(anchor="w")
        lbl_instr1 = ttk.Label(
            card1,
            text="Abre Chrome, inicia sesión y elige tu grupo.\nAl ver las calificaciones, haz clic en 'Detectar Alumnos'.",
            foreground=self.COLOR_MUTED
        )
        lbl_instr1.pack(anchor="w", pady=(3, 8))

        btn_row1 = tk.Frame(card1, bg=self.COLOR_CARD)
        btn_row1.pack(fill=tk.X, pady=4)

        self.btn_open_chrome = ttk.Button(btn_row1, text="🚀 Abrir / Conectar Chrome", style="Primary.TButton", command=self.on_open_chrome)
        self.btn_open_chrome.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_detect_students = ttk.Button(btn_row1, text="🔄 Detectar Alumnos", command=self.on_detect_students)
        self.btn_detect_students.pack(side=tk.LEFT)

        self.lbl_browser_status = ttk.Label(card1, text="⚪ Estado: Chrome no conectado", foreground=self.COLOR_MUTED)
        self.lbl_browser_status.pack(anchor="w", pady=(6, 2))

        self.lbl_students_count = ttk.Label(card1, text="Sin alumnos detectados", font=("Segoe UI", 9, "bold"))
        self.lbl_students_count.pack(anchor="w")

        # Card 2: Archivo Excel y Parcial
        card2 = ttk.Frame(top_panels, style="Card.TFrame", padding=12)
        card2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))

        ttk.Label(card2, text="2. Archivo Excel y Parcial", style="SectionTitle.TLabel").pack(anchor="w")

        file_row = tk.Frame(card2, bg=self.COLOR_CARD)
        file_row.pack(fill=tk.X, pady=(5, 5))

        self.btn_browse_excel = ttk.Button(file_row, text="📂 Seleccionar Excel...", command=self.on_browse_excel)
        self.btn_browse_excel.pack(side=tk.LEFT, padx=(0, 6))

        self.lbl_excel_name = ttk.Label(file_row, text="Ningún archivo seleccionado", foreground=self.COLOR_MUTED)
        self.lbl_excel_name.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Selectores de Parcial, Hoja y Columna
        selectors_frame = tk.Frame(card2, bg=self.COLOR_CARD)
        selectors_frame.pack(fill=tk.X, pady=4)

        # Parcial
        tk.Label(selectors_frame, text="Parcial:", bg=self.COLOR_CARD, font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 4), pady=2)
        self.var_partial = tk.StringVar(value="Parcial 1")
        self.cb_partial = ttk.Combobox(selectors_frame, textvariable=self.var_partial, values=["Parcial 1", "Parcial 2", "Parcial 3"], state="readonly", width=12)
        self.cb_partial.grid(row=0, column=1, sticky="w", padx=(0, 15), pady=2)
        self.cb_partial.bind("<<ComboboxSelected>>", self.on_partial_changed)

        # Hoja
        tk.Label(selectors_frame, text="Pestaña Excel:", bg=self.COLOR_CARD).grid(row=0, column=2, sticky="w", padx=(0, 4), pady=2)
        self.var_sheet = tk.StringVar()
        self.cb_sheet = ttk.Combobox(selectors_frame, textvariable=self.var_sheet, values=[], state="readonly", width=16)
        self.cb_sheet.grid(row=0, column=3, sticky="w", pady=2)
        self.cb_sheet.bind("<<ComboboxSelected>>", self.on_sheet_changed)

        # Columna Total
        tk.Label(selectors_frame, text="Col. Calificación:", bg=self.COLOR_CARD).grid(row=1, column=0, sticky="w", padx=(0, 4), pady=4)
        self.var_total_col = tk.StringVar()
        self.cb_total_col = ttk.Combobox(selectors_frame, textvariable=self.var_total_col, values=[], state="readonly", width=12)
        self.cb_total_col.grid(row=1, column=1, sticky="w", padx=(0, 15), pady=4)
        self.cb_total_col.bind("<<ComboboxSelected>>", self.on_total_col_changed)

        self.lbl_match_summary = ttk.Label(card2, text="", font=("Segoe UI", 8), foreground=self.COLOR_MUTED)
        self.lbl_match_summary.pack(anchor="w", pady=(2, 0))

        # Panel Central: Vista Previa de Tabla de Emparejamiento
        card_table = ttk.Frame(main_container, style="Card.TFrame", padding=12)
        card_table.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        ttk.Label(card_table, text="3. Vista Previa de Asignación (Correspondencia Alumnos vs Calificaciones)", style="SectionTitle.TLabel").pack(anchor="w", pady=(0, 6))

        # Treeview con Scrollbar
        table_container = tk.Frame(card_table, bg=self.COLOR_CARD)
        table_container.pack(fill=tk.BOTH, expand=True)

        columns = ("cuenta", "nombre_upp", "nombre_excel", "calificacion", "match")
        self.tree = ttk.Treeview(table_container, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("cuenta", text="Cuenta UPP")
        self.tree.heading("nombre_upp", text="Nombre en SIIUPP")
        self.tree.heading("nombre_excel", text="Nombre en Excel")
        self.tree.heading("calificacion", text="Nota a Pegar")
        self.tree.heading("match", text="Coincidencia")

        self.tree.column("cuenta", width=110, anchor="center")
        self.tree.column("nombre_upp", width=280, anchor="w")
        self.tree.column("nombre_excel", width=280, anchor="w")
        self.tree.column("calificacion", width=100, anchor="center")
        self.tree.column("match", width=120, anchor="center")

        tree_scroll = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Tags de color para el treeview
        self.tree.tag_configure("ok", background="#ECFDF5")         # verde claro
        self.tree.tag_configure("missing", background="#FEF2F2")    # rojo claro
        self.tree.tag_configure("empty", background="#F9FAFB")

        # Panel Inferior: Acción de Pegar y Aviso de Seguridad
        bottom_frame = tk.Frame(main_container, bg=self.COLOR_BG)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)

        notice_frame = tk.Frame(bottom_frame, bg="#EFF6FF", highlightbackground="#BFDBFE", highlightthickness=1, padx=10, pady=6)
        notice_frame.pack(fill=tk.X, pady=(0, 8))

        lbl_notice = tk.Label(
            notice_frame,
            text="🔒 Modo Seguro Activo: La aplicación ÚNICAMENTE pegará las calificaciones en la columna del parcial seleccionado.\nNo dará clic en guardar, enviar ni alterará asistencias ni otros campos. Tú revisas y guardas manualmente en Chrome.",
            bg="#EFF6FF",
            fg="#1E3A8A",
            font=("Segoe UI", 8),
            justify="left"
        )
        lbl_notice.pack(anchor="w")

        action_row = tk.Frame(bottom_frame, bg=self.COLOR_BG)
        action_row.pack(fill=tk.X)

        self.btn_paste = ttk.Button(
            action_row,
            text="✍️  PEGAR CALIFICACIONES EN CHROME",
            style="Success.TButton",
            command=self.on_paste_grades
        )
        self.btn_paste.pack(side=tk.RIGHT)

        self.lbl_action_status = ttk.Label(action_row, text="", font=("Segoe UI", 9, "bold"))
        self.lbl_action_status.pack(side=tk.LEFT, pady=4)

        # Conectar automáticamente si Chrome ya estaba en el puerto 9222
        self.root.after(300, self._auto_connect_existing_chrome)

    def _auto_connect_existing_chrome(self):
        """Si Chrome ya está abierto con depuración remota, conectarse silenciosamente."""
        ok, msg = self.browser_manager.connect_or_launch(9222)
        if ok and "127.0.0.1:9222" in msg:
            self.lbl_browser_status.config(text="🟢 Chrome conectado (Sesión activa)", foreground=self.COLOR_SUCCESS)
            # Intentar detectar alumnos si ya está en la página
            self.on_detect_students(silent=True)

    def on_open_chrome(self):
        """Abre o se conecta a Chrome en segundo plano."""
        def run():
            self.lbl_browser_status.config(text="🟡 Conectando con Chrome...", foreground="#D97706")
            ok, msg = self.browser_manager.connect_or_launch()
            if ok:
                self.lbl_browser_status.config(text="🟢 Chrome conectado", foreground=self.COLOR_SUCCESS)
                messagebox.showinfo(
                    "Chrome Listo",
                    f"{msg}\n\n1. Inicia sesión en el portal.\n2. Elige tu grupo y dirígete a Calificaciones.\n3. Regresa aquí y presiona 'Detectar Alumnos'."
                )
            else:
                self.lbl_browser_status.config(text="🔴 Error al conectar Chrome", foreground="#DC2626")
                messagebox.showerror("Error de Chrome", msg)

        threading.Thread(target=run, daemon=True).start()

    def on_detect_students(self, silent: bool = False):
        """Extrae los alumnos de la página actual de calificaciones en Chrome."""
        if not self.browser_manager.is_connected():
            if not silent:
                messagebox.showwarning("Navegador no conectado", "Primero presiona 'Abrir / Conectar Chrome'.")
            return

        students = self.browser_manager.get_students()
        if not students:
            self.lbl_students_count.config(text="⚠️ No se detectó la tabla de alumnos.", foreground="#D97706")
            if not silent:
                messagebox.showwarning(
                    "Tabla no encontrada",
                    "Asegúrate de tener abierta la página de captura de calificaciones del grupo en Chrome (vCalificacionesVer2.php)."
                )
            return

        self.upp_students = students
        self.lbl_students_count.config(
            text=f"✅ {len(students)} alumnos detectados en el grupo.",
            foreground=self.COLOR_SUCCESS
        )

        # Si ya hay un Excel cargado, recalcular el emparejamiento
        if self.df_excel is not None:
            self._update_matching_preview()
        else:
            # Mostrar la lista de alumnos en el treeview sin notas aún
            self._populate_treeview([
                {
                    "cuenta": s["cuenta"],
                    "nombre_upp": s["nombre"],
                    "nombre_excel": "-",
                    "calificacion": "-",
                    "metodo_match": "Pendiente de Excel",
                    "encontrado": False
                }
                for s in self.upp_students
            ])

        if not silent:
            messagebox.showinfo("Alumnos Detectados", f"Se detectaron con éxito {len(students)} alumnos en el grupo de Chrome.")

    def on_browse_excel(self):
        """Selecciona el archivo Excel y carga las pestañas."""
        path = filedialog.askopenfilename(
            title="Seleccionar archivo de Excel",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls *.xlsm"), ("Todos los archivos", "*.*")]
        )
        if not path:
            return

        self.excel_file_path = path
        filename = os.path.basename(path)
        self.lbl_excel_name.config(text=filename, foreground=self.COLOR_TEXT)

        try:
            sheet_names = ExcelManager.get_sheet_names(path)
            self.cb_sheet["values"] = sheet_names

            # Intentar seleccionar la hoja adecuada según el parcial actual
            partial_num = self._get_selected_partial_num()
            best_sheet = ExcelManager.find_best_sheet(sheet_names, partial_num)
            self.var_sheet.set(best_sheet)

            # Cargar datos de la hoja
            self._load_sheet_content(best_sheet)
        except Exception as e:
            messagebox.showerror("Error al leer Excel", f"No se pudo leer el archivo:\n{e}")

    def _get_selected_partial_num(self) -> int:
        val = self.var_partial.get()
        if "1" in val:
            return 1
        elif "2" in val:
            return 2
        elif "3" in val:
            return 3
        return 1

    def on_partial_changed(self, event=None):
        """Al cambiar de parcial (ej. de Parcial 1 a Parcial 2), sugerir la hoja correspondiente."""
        if not self.excel_file_path or not self.cb_sheet["values"]:
            return
        partial_num = self._get_selected_partial_num()
        best_sheet = ExcelManager.find_best_sheet(list(self.cb_sheet["values"]), partial_num)
        if best_sheet and best_sheet != self.var_sheet.get():
            self.var_sheet.set(best_sheet)
            self._load_sheet_content(best_sheet)
        else:
            self._update_matching_preview()

    def on_sheet_changed(self, event=None):
        sheet = self.var_sheet.get()
        if sheet and self.excel_file_path:
            self._load_sheet_content(sheet)

    def _load_sheet_content(self, sheet_name: str):
        try:
            self.df_excel = ExcelManager.load_sheet_data(self.excel_file_path, sheet_name)
            cols = list(self.df_excel.columns)
            self.cb_total_col["values"] = cols

            # Autodetectar columna total
            best_total = ExcelManager.find_total_column(cols)
            self.var_total_col.set(best_total)

            self._update_matching_preview()
        except Exception as e:
            messagebox.showerror("Error al procesar hoja", f"Error cargando la pestaña '{sheet_name}':\n{e}")

    def on_total_col_changed(self, event=None):
        self._update_matching_preview()

    def _update_matching_preview(self):
        """Realiza el emparejamiento entre los alumnos de UPP y el Excel cargado."""
        if self.df_excel is None:
            return

        total_col = self.var_total_col.get()
        if not total_col or total_col not in self.df_excel.columns:
            return

        if not self.upp_students:
            # Si aún no se han detectado alumnos en Chrome, mostrar las filas del Excel directamente
            auto_c, auto_n = ExcelManager.find_identity_columns(list(self.df_excel.columns))
            preview_rows = []
            for idx, r in self.df_excel.head(30).iterrows():
                preview_rows.append({
                    "cuenta": clean_matricula(r.get(auto_c, "")) if auto_c else "",
                    "nombre_upp": "— (Detecta alumnos en Chrome)",
                    "nombre_excel": str(r.get(auto_n, r.iloc[0] if len(r) > 0 else "")),
                    "calificacion": r.get(total_col, ""),
                    "metodo_match": "Vista previa Excel",
                    "encontrado": True
                })
            self._populate_treeview(preview_rows)
            self.lbl_match_summary.config(text=f"Hoja cargada: {len(self.df_excel)} filas en Excel. Presiona 'Detectar Alumnos' para emparejar.")
            return

        self.matched_data = ExcelManager.match_with_upp_students(
            upp_students=self.upp_students,
            df_excel=self.df_excel,
            total_col=total_col
        )

        matched_count = sum(1 for m in self.matched_data if m["encontrado"] and m["calificacion"] != "")
        missing_count = len(self.matched_data) - matched_count

        self.lbl_match_summary.config(
            text=f"Total: {len(self.matched_data)} alumnos | ✅ {matched_count} con calificación | ⚠️ {missing_count} pendientes"
        )
        self._populate_treeview(self.matched_data)

    def _populate_treeview(self, data: list[dict]):
        self.tree.delete(*self.tree.get_children())
        for item in data:
            val_grade = item.get("calificacion", "")
            if val_grade is None:
                val_grade = ""
            tag = "ok" if item.get("encontrado", False) and str(val_grade).strip() != "" else "missing"

            self.tree.insert(
                "",
                tk.END,
                values=(
                    item.get("cuenta", ""),
                    item.get("nombre_upp", ""),
                    item.get("nombre_excel", ""),
                    str(val_grade),
                    item.get("metodo_match", "")
                ),
                tags=(tag,)
            )

    def on_paste_grades(self):
        """Pega las calificaciones en la página de Chrome."""
        if not self.browser_manager.is_connected():
            messagebox.showwarning("Navegador no conectado", "Chrome no está conectado.")
            return

        if not self.upp_students:
            messagebox.showwarning("Alumnos no detectados", "Primero haz clic en 'Detectar Alumnos'.")
            return

        if not self.matched_data:
            messagebox.showwarning("Sin datos", "Selecciona un archivo de Excel y verifica la columna de notas.")
            return

        partial_idx = self._get_selected_partial_num() - 1  # 0 para Parcial 1, 1 para Parcial 2, 2 para Parcial 3
        partial_name = self.var_partial.get()

        # Filtrar calificaciones válidas
        grades_to_paste = {}
        for item in self.matched_data:
            grade = item.get("calificacion")
            if grade is not None and str(grade).strip() != "" and str(grade).strip() != "-":
                grades_to_paste[item["cuenta"]] = grade

        if not grades_to_paste:
            messagebox.showwarning("Sin calificaciones", "No se encontraron calificaciones válidas para pegar.")
            return

        confirm = messagebox.askyesno(
            "Confirmar Asignación",
            f"¿Deseas pegar {len(grades_to_paste)} calificaciones en la columna '{partial_name}' de Chrome?\n\n"
            f"Nota: Solo se escribirán las notas. No se dará clic a guardar ni enviar."
        )
        if not confirm:
            return

        self.lbl_action_status.config(text="⏳ Pegando calificaciones en Chrome...", foreground="#D97706")
        self.root.update()

        pasted_count, errors = self.browser_manager.paste_grades(grades_to_paste, partial_idx)

        if errors:
            msg_err = "\n".join(errors[:5])
            self.lbl_action_status.config(
                text=f"⚠️ {pasted_count} calificaciones pegadas con {len(errors)} advertencias.",
                foreground="#D97706"
            )
            messagebox.showwarning(
                "Pegado con advertencias",
                f"Se pegaron {pasted_count} calificaciones.\n\nDetalles de advertencias:\n{msg_err}"
            )
        else:
            self.lbl_action_status.config(
                text=f"🎉 ¡Éxito! Se pegaron las {pasted_count} calificaciones en {partial_name}.",
                foreground=self.COLOR_SUCCESS
            )
            messagebox.showinfo(
                "¡Calificaciones Pegadas!",
                f"✅ Se pegaron con éxito las {pasted_count} calificaciones en la columna '{partial_name}'.\n\n"
                f"Por favor, revisa en tu pantalla de Chrome que todo se vea correcto y haz clic en 'Guardar' cuando lo consideres conveniente."
            )


def main():
    root = tk.Tk()
    app = SIIUPPApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

