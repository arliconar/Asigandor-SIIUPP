"""
Módulo central (core) para la interacción con Chrome (SIIUPP) y análisis de archivos Excel.
Usa únicamente librerías estándar y openpyxl (puro Python, sin DLLs C de NumPy/Pandas).
"""

import re
import unicodedata
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By


def normalize_text(text: str) -> str:
    """Normaliza texto: mayúsculas, sin acentos ni espacios repetidos."""
    if not text:
        return ""
    text = str(text).strip().upper()
    # Eliminar diacríticos/acentos
    nfkd = unicodedata.normalize("NFKD", text)
    cleaned = "".join([c for c in nfkd if not unicodedata.combining(c)])
    # Limpiar espacios
    return re.sub(r"\s+", " ", cleaned)


def clean_matricula(val) -> str:
    """Extrae dígitos de una matrícula."""
    if val is None:
        return ""
    digits = re.sub(r"\D", "", str(val).strip())
    return digits


class BrowserManager:
    """Controlador de Selenium para Google Chrome y la plataforma SIIUPP."""

    UPP_LOGIN_URL = "https://sii.upp.edu.mx/vista/academicos/vLoginAcademicos.php"

    def __init__(self):
        self.driver = None

    def connect_or_launch(self, remote_port: int = 9222) -> tuple[bool, str]:
        """
        Intenta conectarse a un Chrome ya abierto con puerto de depuración.
        Si no está disponible, lanza una nueva instancia de Chrome y navega al login.
        """
        # 1. Intentar conectar a Chrome existente en el puerto
        try:
            options = Options()
            options.add_experimental_option("debuggerAddress", f"127.0.0.1:{remote_port}")
            test_driver = webdriver.Chrome(options=options)
            _ = test_driver.title
            self.driver = test_driver
            return True, f"Conectado a Chrome existente en 127.0.0.1:{remote_port}"
        except Exception:
            pass

        # 2. Si no se pudo conectar, abrir una nueva instancia de Chrome
        try:
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)

            self.driver = webdriver.Chrome(options=options)
            self.driver.get(self.UPP_LOGIN_URL)
            return True, "Nueva ventana de Chrome abierta en el portal de UPP."
        except Exception as e:
            err_msg = str(e)
            if "binary_location" in err_msg.lower() or "cannot find chrome" in err_msg.lower():
                return False, "No se encontró Google Chrome instalado en tu computadora. Asegúrate de tener instalado Chrome estándar."
            return False, f"No se pudo iniciar Chrome: {err_msg}"

    def is_connected(self) -> bool:
        if self.driver is None:
            return False
        try:
            _ = self.driver.title
            return True
        except Exception:
            self.driver = None
            return False

    def get_current_url(self) -> str:
        if not self.is_connected():
            return ""
        try:
            return self.driver.current_url
        except Exception:
            return ""

    def get_students(self) -> list[dict]:
        """
        Extrae dinámicamente la lista de alumnos actualmente cargada en la tabla de vCalificacionesVer2.php.
        Retorna: [{'cuenta': '2631127239', 'nombre': 'AMARO RAMIREZ NATALY ANAHI'}, ...]
        """
        if not self.is_connected():
            return []

        students = []
        try:
            rows = self.driver.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                tds = row.find_elements(By.TAG_NAME, "td")
                if len(tds) >= 2:
                    cuenta = clean_matricula(tds[0].text)
                    nombre = tds[1].text.strip()
                    # Las matrículas en UPP son numéricas de al menos 8 dígitos
                    if len(cuenta) >= 8 and len(nombre) > 3:
                        students.append({
                            "cuenta": cuenta,
                            "nombre": nombre,
                            "nombre_norm": normalize_text(nombre)
                        })
        except Exception as e:
            print(f"Error al extraer alumnos: {e}")

        return students

    def paste_grades(self, student_grades: dict[str, float | str], partial_index: int) -> tuple[int, list[str]]:
        """
        Pega las calificaciones en los campos 'calf{partial_index}{cuenta}'.
        partial_index: 0 = Parcial 1, 1 = Parcial 2, 2 = Parcial 3.
        NO envía formulario ni hace clic en guardar.
        Retorna: (total_pegados, lista_de_errores)
        """
        if not self.is_connected():
            return 0, ["El navegador no está conectado."]

        field_prefix = f"calf{partial_index}"
        pasted_count = 0
        errors = []

        js_setter = """
            const inputId = arguments[0];
            const val = arguments[1];
            const inp = document.getElementById(inputId);
            if (!inp) {
                return 'NOT_FOUND';
            }
            inp.value = val;
            inp.dispatchEvent(new Event('input', { bubbles: true }));
            inp.dispatchEvent(new Event('change', { bubbles: true }));
            return 'OK';
        """

        for cuenta, grade in student_grades.items():
            if grade is None or str(grade).strip() == "":
                continue

            input_id = f"{field_prefix}{cuenta}"
            try:
                # Formatear calificación (restringir a 0.0 - 10.0 con 1 decimal o entero)
                grade_num = float(str(grade).replace(",", "."))
                grade_str = f"{grade_num:.1f}" if grade_num % 1 != 0 else f"{int(grade_num)}"
            except ValueError:
                grade_str = str(grade).strip()

            try:
                result = self.driver.execute_script(js_setter, input_id, grade_str)
                if result == "OK":
                    pasted_count += 1
                else:
                    errors.append(f"No se encontró el campo {input_id} para la cuenta {cuenta}")
            except Exception as e:
                errors.append(f"Error asignando {cuenta}: {e}")

        return pasted_count, errors


class ExcelTable:
    """Estructura de datos ligera y 100% en Python puro para representar tablas de Excel."""

    def __init__(self, columns: list[str], rows: list[dict]):
        self.columns = columns
        self.rows = rows

    def __len__(self):
        return len(self.rows)

    def iterrows(self):
        for idx, row in enumerate(self.rows):
            yield idx, row

    def head(self, n=30):
        return ExcelTable(self.columns, self.rows[:n])


class ExcelManager:
    """Lector e interpretador de archivos Excel usando openpyxl (sin dependencias C/NumPy)."""

    @staticmethod
    def get_sheet_names(file_path: str) -> list[str]:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names

    @staticmethod
    def find_best_sheet(sheet_names: list[str], partial_num: int) -> str:
        """Busca la pestaña más probable para el parcial indicado (ej. Parcial 1)."""
        search_terms = [
            f"parcial {partial_num}",
            f"parcial{partial_num}",
            f"p{partial_num}",
            f"{partial_num}er parcial",
            f"{partial_num}do parcial",
            f"{partial_num}to parcial",
            f"{partial_num} parcial",
        ]
        norm_names = {s: normalize_text(s) for s in sheet_names}

        for sheet, norm in norm_names.items():
            for term in search_terms:
                if term in norm.lower():
                    return sheet

        if 0 <= (partial_num - 1) < len(sheet_names):
            return sheet_names[partial_num - 1]
        return sheet_names[0] if sheet_names else ""

    @staticmethod
    def load_sheet_data(file_path: str, sheet_name: str) -> ExcelTable:
        """Carga la hoja detectando automáticamente la fila de encabezados."""
        wb = openpyxl.load_workbook(file_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            ws = wb.active
        else:
            ws = wb[sheet_name]

        all_rows = []
        for row in ws.iter_rows(values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                all_rows.append(list(row))
        wb.close()

        if not all_rows:
            return ExcelTable([], [])

        # Buscar en las primeras 15 filas cuál contiene palabras clave de encabezados
        header_idx = 0
        keywords = ["NOMBRE", "ALUMNO", "MATRICULA", "CUENTA", "TOTAL", "CALIFICACION", "CALIF"]

        for idx, r in enumerate(all_rows[:15]):
            row_str = " ".join([str(val).upper() for val in r if val is not None])
            matches = sum(1 for kw in keywords if kw in row_str)
            if matches >= 2:
                header_idx = idx
                break

        header_raw = all_rows[header_idx]
        columns = []
        seen = {}
        for i, c in enumerate(header_raw):
            col_name = str(c).strip() if c is not None and str(c).strip() != "" else f"Columna_{i+1}"
            if col_name in seen:
                seen[col_name] += 1
                col_name = f"{col_name}_{seen[col_name]}"
            else:
                seen[col_name] = 1
            columns.append(col_name)

        data_rows = []
        for r in all_rows[header_idx + 1:]:
            if not any(val is not None and str(val).strip() != "" for val in r):
                continue
            row_dict = {}
            for col_name, val in zip(columns, r):
                row_dict[col_name] = val
            data_rows.append(row_dict)

        return ExcelTable(columns, data_rows)

    @staticmethod
    def find_total_column(columns: list[str]) -> str:
        """Encuentra la columna correspondiente al Total o calificación final del parcial."""
        for col in columns:
            norm = normalize_text(col)
            if norm == "TOTAL":
                return col

        for col in columns:
            norm = normalize_text(col)
            if "TOTAL" in norm:
                return col

        for col in columns:
            norm = normalize_text(col)
            if any(k in norm for k in ["CALIFICACION", "CALIF", "PROMEDIO"]):
                return col

        return columns[-1] if columns else ""

    @staticmethod
    def find_identity_columns(columns: list[str]) -> tuple[str | None, str | None]:
        """Detecta la columna de matrícula/cuenta y la de nombre."""
        col_cuenta = None
        col_nombre = None

        for col in columns:
            norm = normalize_text(col)
            if any(k in norm for k in ["CUENTA", "MATRICULA", "ID", "EXPEDIENTE"]):
                if not col_cuenta:
                    col_cuenta = col
            if any(k in norm for k in ["NOMBRE", "ALUMNO", "ESTUDIANTE"]):
                if not col_nombre:
                    col_nombre = col

        return col_cuenta, col_nombre

    @classmethod
    def match_with_upp_students(
        cls,
        upp_students: list[dict],
        df_excel: ExcelTable,
        total_col: str,
        cuenta_col: str | None = None,
        nombre_col: str | None = None,
    ) -> list[dict]:
        """
        Empareja alumnos de UPP con las filas del Excel.
        Retorna lista de diccionarios con el resultado del emparejamiento.
        """
        auto_cuenta, auto_nombre = cls.find_identity_columns(list(df_excel.columns))
        col_cuenta = cuenta_col or auto_cuenta
        col_nombre = nombre_col or auto_nombre

        # Preparar registros de Excel
        excel_records = []
        for idx, row in df_excel.iterrows():
            grade_raw = row.get(total_col, None)
            if grade_raw is None or str(grade_raw).strip() == "" or str(grade_raw).lower() == "nan":
                grade_val = ""
            else:
                grade_val = grade_raw

            c_val = clean_matricula(row.get(col_cuenta, "")) if col_cuenta else ""
            n_val = str(row.get(col_nombre, "")).strip() if col_nombre else ""

            excel_records.append({
                "index": idx,
                "cuenta": c_val,
                "nombre": n_val,
                "nombre_norm": normalize_text(n_val),
                "grade": grade_val
            })

        matched_results = []

        excel_by_cuenta = {r["cuenta"]: r for r in excel_records if r["cuenta"]}
        excel_by_name = {r["nombre_norm"]: r for r in excel_records if r["nombre_norm"]}

        for upp_idx, student in enumerate(upp_students):
            s_cuenta = student["cuenta"]
            s_norm = student["nombre_norm"]

            match_record = None
            match_method = "No encontrado"

            # 1. Emparejar por cuenta/matrícula
            if s_cuenta in excel_by_cuenta:
                match_record = excel_by_cuenta[s_cuenta]
                match_method = "Matrícula"
            # 2. Emparejar por nombre exacto normalizado
            elif s_norm in excel_by_name:
                match_record = excel_by_name[s_norm]
                match_method = "Nombre exacto"
            else:
                # 3. Emparejar por inclusión parcial de nombre
                for e_norm, rec in excel_by_name.items():
                    if s_norm in e_norm or e_norm in s_norm:
                        match_record = rec
                        match_method = "Nombre similar"
                        break

            # 4. Respaldo por orden de fila si coincide la cantidad
            if not match_record and upp_idx < len(excel_records):
                if not col_cuenta and not col_nombre:
                    match_record = excel_records[upp_idx]
                    match_method = "Orden de lista"

            grade_to_assign = match_record["grade"] if match_record else ""
            matched_results.append({
                "cuenta": s_cuenta,
                "nombre_upp": student["nombre"],
                "nombre_excel": match_record["nombre"] if match_record else "-",
                "calificacion": grade_to_assign,
                "metodo_match": match_method,
                "encontrado": match_record is not None
            })

        return matched_results
