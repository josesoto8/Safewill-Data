#Librerías.

#Librería para conectar el código con el sistema operativo.
import os
#Librería para poder usar archivos .csv
import csv
#Librería para la función de guardar la carpeta donde se guardaran los datos.
import json
#Librería para guardar archivos dentro del código.
import sys
#Librería para crear una interfaz grafica.
import tkinter as tk
from tkinter import filedialog, messagebox
#Librería para leer y modificar archivos de Excel.
import openpyxl
#Librería para trabajar con horas y fechas.
from datetime import datetime, timedelta
#Librería para modificar aún mas la interfaz.
import customtkinter as ctk
import getpass
import ctypes
#Librería con la funcion de aleatoriedad que se usa en los saludos.
import random

def obtener_nombre_usuario():
    #Para obtener el nombre visible de la cuenta que usa el sistema operativo.
    try:
        GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
        NameDisplay = 3
        size = ctypes.c_ulong(0)
        GetUserNameEx(NameDisplay, None, ctypes.byref(size))
        buffer = ctypes.create_unicode_buffer(size.value)
        if GetUserNameEx(NameDisplay, buffer, ctypes.byref(size)) and buffer.value.strip():
            # Retorna el primer nombre de la cuenta
            return buffer.value.split()[0].capitalize()
    except Exception:
        pass

    # 2. Si falla, busca el usuario en las variables de entorno
    nombre = os.environ.get("USERNAME") or os.environ.get("USER")
    
    if not nombre:
        try:
            nombre = os.getlogin()
        except Exception:
            try:
                nombre = getpass.getuser()
            except Exception:
                nombre = "Usuario"

    return nombre.capitalize()

#Configuración visual de CustomTkinter
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

#Asigna un ID de aplicación en Windows para el icono de la barra de tareas
try:
    import ctypes
    myappid = 'safewill.data.v1.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

#Función para obtener la ruta del icono del .exe
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

#Para el redimensionamiento de la aplicación
class CSVFixerAndMergerApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Safewill Data")
        
        #Cambio de tamaño de la ventana inicial y la opcion de poder maximizar la ventana
        self.root.geometry("620x600")
        self.root.minsize(580, 560)  #Establece un tamaño mínimo para no deformar controles
        self.root.resizable(True, True)  #Permite modificar el tamaño de la ventana
        # ------------------------------------------

        self.root.configure(fg_color="white")

        #Icono para Windows
        icono_path = resource_path("icono.ico")

        try:
            if os.path.exists(icono_path):
                self.root.iconbitmap(default=icono_path)
                self.root.iconbitmap(icono_path)
            elif getattr(sys, "frozen", False):
                self.root.iconbitmap(default=sys.executable)
                self.root.iconbitmap(sys.executable)
        except Exception:
            pass

        self.file_paths = []
        self.cleaned_files_data = []
        self.output_folder = self._cargar_configuracion()

        self.create_widgets()
        self._actualizar_lbl_folder()
        self._actualizar_estado_botones()

    #Función de guardado en JSON para guardar la configuración de ruta de guardado de la carpeta
    def _get_config_path(self):
        """Retorna la ruta del archivo de configuración JSON."""
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, "config.json")

    def _cargar_configuracion(self):
        """Carga la carpeta guardada previamente si existe."""
        config_path = self._get_config_path()
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    folder = data.get("output_folder", "")
                    if os.path.exists(folder):
                        return folder
            except Exception:
                pass
        return ""

    def _guardar_configuracion(self):
        """Guarda la carpeta seleccionada en el archivo JSON."""
        config_path = self._get_config_path()
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump({"output_folder": self.output_folder}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    #Interfaz, saludo inicial y botones
    #Interfaz, titulos y botones
    def create_widgets(self):
        nombre_usuario = obtener_nombre_usuario()

        mensajes_bienvenida = [
            (f"¡Hola {nombre_usuario}!", "¿Qué archivos quieres que unamos hoy?"),
            (f"¡Bienvenido, {nombre_usuario}!", "Vamos a unir esos CSVs."),
            (f"¡Qué tal, {nombre_usuario}!", "¿Listo para procesar tus datos?"),
            ("HOY NO QUIERO AYUDARTE!", f"Mentira, {nombre_usuario} a ver esos datos."),
            ("¡Buen día!", f"Es hora de trabajar, {nombre_usuario}."),
            ("¡Hola de nuevo!", "Selecciona los datos y yo hago el resto."),
            (f"¡ahhhh, no dormi bien!", "aun asi estoy aqui para ayudartzZZzzZ."),
            ('print("Hello World!")', ".py"),
            (f"00010011100110110101:", f"Tranquilo {nombre_usuario}, hoy no habrá errores.")
        ]

        titulo_elegido, subtitulo_elegido = random.choice(mensajes_bienvenida)

        self.lbl_titulo = ctk.CTkLabel(
            self.root,
            text=titulo_elegido,
            font=("Arial", 25, "bold")
        )
        self.lbl_titulo.pack(pady=(15, 2))

        self.lbl_subtitulo = ctk.CTkLabel(
            self.root,
            text=subtitulo_elegido,
            font=("Arial", 13, "italic"),
            text_color="gray"
        )
        self.lbl_subtitulo.pack(pady=(0, 15))

        #Cargar archivos
        file_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        file_frame.pack(fill="x", padx=20, pady=5)

        self.btn_select = ctk.CTkButton(
            file_frame,
            text="Cargar archivos CSV",
            command=self.load_csvs,
            #Colores de los botones
            fg_color="#FFCE44",
            hover_color="#FFCE44",
            text_color="#000000",
            font=("Arial", 14, "bold"),
            corner_radius=24,
            height=36,
        )
        self.btn_select.pack(side="left")

        self.lbl_file = ctk.CTkLabel(
            file_frame,
            text="Sin archivos seleccionados",
            text_color="gray",
            anchor="w",
            wraplength=350,
            justify="left",
        )
        self.lbl_file.pack(side="left", fill="x", expand=True, padx=10)

        #Seleccionar carpeta de destino
        folder_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        folder_frame.pack(fill="x", padx=20, pady=5)

        self.btn_select_folder = ctk.CTkButton(
            folder_frame,
            text="Carpeta de destino",
            command=self.select_output_folder,
            fg_color="#007AFF",
            hover_color="#007AFF",
            text_color="#F5F3F3",
            font=("Arial", 14, "bold"),
            corner_radius=24,
            height=36,
        )
        self.btn_select_folder.pack(side="left")

        self.lbl_folder = ctk.CTkLabel(
            folder_frame,
            text="Sin carpeta seleccionada",
            text_color="gray",
            anchor="w",
            wraplength=350,
            justify="left",
        )
        self.lbl_folder.pack(side="left", fill="x", expand=True, padx=10)

        info_frame = ctk.CTkFrame(self.root, corner_radius=10)
        info_frame.pack(fill="x", padx=20, pady=12)

        ctk.CTkLabel(
            info_frame,
            text="Uso",
            font=("Arial", 12, "bold"),
        ).pack(anchor="w", padx=10, pady=(8, 2))

        self.lbl_info = ctk.CTkLabel(
            info_frame,
            text="• Selecciona los archivos .csv y la carpeta de destino.\n"
                 "• Elige si quieres los datos completos o promediados.\n"
                 "• Guarda el archivo consolidado en formato .csv o .xlsx.",
            font=("Arial", 12),
            justify="left",
        )
        self.lbl_info.pack(anchor="w", padx=10, pady=(0, 8))

        #Promedio
        self.var_promediar = tk.BooleanVar(value=False)
        self.chk_promedio = ctk.CTkCheckBox(
            self.root,
            text="Promediar datos por hora (ej. Hora 12, Hora 13...)",
            variable=self.var_promediar,
            font=("Arial", 15),
            checkbox_height=20,
            checkbox_width=20,
            corner_radius=6,
        )
        self.chk_promedio.pack(pady=10)

        #Botones para guardar
        btn_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)

        self.btn_save_csv = ctk.CTkButton(
            btn_frame,
            text="Guardar como CSV (.csv)",
            command=self.save_csv,
            fg_color="#022B8B",
            hover_color="#002A8B",
            font=("Arial", 15, "bold"),
            corner_radius=24,
            height=40,
            state="disabled",
        )
        self.btn_save_csv.pack(side="left", expand=True, fill="x", padx=5)

        self.btn_save_excel = ctk.CTkButton(
            btn_frame,
            text="Guardar como Excel (.xlsx)",
            command=self.save_excel,
            fg_color="#F83B3B",
            hover_color="#E24343",
            text_color="#000000",
            font=("Arial", 15, "bold"),
            corner_radius=24,
            height=40,
            state="disabled",
        )
        self.btn_save_excel.pack(side="left", expand=True, fill="x", padx=5)

        self.lbl_status = ctk.CTkLabel(
            self.root, text="", text_color="green", font=("Arial", 11, "italic")
        )
        self.lbl_status.pack(pady=5)

    def select_output_folder(self):
        initial_dir = self.output_folder if self.output_folder and os.path.exists(self.output_folder) else os.getcwd()
        folder = filedialog.askdirectory(
            title="Selecciona la carpeta donde se guardarán los archivos",
            initialdir=initial_dir
        )
        if folder:
            self.output_folder = folder
            self._guardar_configuracion()
            self._actualizar_lbl_folder()
            self._actualizar_estado_botones()

    def _actualizar_lbl_folder(self):
        if self.output_folder and os.path.exists(self.output_folder):
            self.lbl_folder.configure(text=f"Carpeta: {self.output_folder}", text_color="black")
        else:
            self.lbl_folder.configure(text="Selecciona una carpeta para guardar", text_color="#E43F43")

    def _actualizar_estado_botones(self):
        listo = bool(self.cleaned_files_data) and bool(self.output_folder) and os.path.exists(self.output_folder)
        state = "normal" if listo else "disabled"
        self.btn_save_csv.configure(state=state)
        self.btn_save_excel.configure(state=state)

    def _parse_and_clean_value(self, val):
        """Corrige fechas YY-M-D o YYYY-M-D a DD/MM/YYYY y convierte valores numéricos."""
        val = str(val).strip()

        subparts = val.split("-")
        if len(subparts) == 3 and all(p.isdigit() for p in subparts):
            yy, mm, dd = subparts
            yyyy = f"20{yy}" if len(yy) == 2 else yy
            return f"{dd.zfill(2)}/{mm.zfill(2)}/{yyyy}"

        try:
            if "." in val:
                return float(val)
            return int(val)
        except ValueError:
            return val

    def _leer_y_corregir_csv(self, filepath):
        filas = []
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            reader = csv.reader(f, delimiter=",")
            for row in reader:
                if not row or any("sep=" in str(celda).lower() for celda in row):
                    continue

                parsed_row = [self._parse_and_clean_value(celda) for celda in row]
                filas.append(parsed_row)
        return filas

    def load_csvs(self):
        files_selected = filedialog.askopenfilenames(
            title="Selecciona los archivos CSV a corregir y unir",
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
        )
        if not files_selected:
            return

        self.file_paths = files_selected
        nombres = [os.path.basename(f) for f in self.file_paths]
        texto_archivos = f"{len(nombres)} archivos cargados."
        self.lbl_file.configure(text=texto_archivos, text_color="black")

        try:
            self.cleaned_files_data = [
                self._leer_y_corregir_csv(f) for f in self.file_paths
            ]

            if not any(self.cleaned_files_data):
                messagebox.showerror("Error", "Los archivos seleccionados están vacíos.")
                return

            self._actualizar_estado_botones()
            self.lbl_status.configure(
                text="¡Archivos cargados correctamente!", text_color="#2196F3"
            )

        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron procesar los archivos:\n{e}")

    def _extraer_hora_simplificada(self, celda_time):
        """Extrae la hora pura para el cálculo del promedio (ej. '11:09:34' -> '11')."""
        texto = str(celda_time).strip()
        if ":" in texto:
            hora_str = texto.split(":")[0]
            if hora_str.isdigit():
                return str(int(hora_str))
        return texto

    def _parse_dt(self, fecha_str, hora_str):
        """Convierte (fecha, hora) a un objeto datetime para comparaciones precisas."""
        f_str = str(fecha_str).strip()
        dt_date = None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                dt_date = datetime.strptime(f_str, fmt)
                break
            except ValueError:
                pass

        if not dt_date:
            for sep in ("/", "-"):
                if sep in f_str:
                    parts = f_str.split(sep)
                    if len(parts) == 3 and all(p.isdigit() for p in parts):
                        p1, p2, p3 = int(parts[0]), int(parts[1]), int(parts[2])
                        try:
                            if p3 > 1000:
                                dt_date = datetime(p3, p2, p1)
                            elif p1 > 1000:
                                dt_date = datetime(p1, p2, p3)
                        except ValueError:
                            pass
                    break

        if not dt_date:
            return None

        h_str = str(hora_str).strip()
        hour, minute, second = 0, 0, 0
        if ":" in h_str:
            parts = h_str.split(":")
            try:
                hour = int(parts[0])
                minute = int(parts[1]) if len(parts) > 1 else 0
                second = int(parts[2]) if len(parts) > 2 else 0
            except ValueError:
                return None
        else:
            try:
                hour = int(h_str)
            except ValueError:
                return None

        return datetime(dt_date.year, dt_date.month, dt_date.day, hour, minute, second)

    def _procesar_archivo_individual(self, filas, promediar=False):
        """Extrae el tipo de sensor, la unidad original y las filas de datos."""
        if not filas:
            return "", "Value", []

        nombre_sensor = ""
        if len(filas[0]) >= 2:
            nombre_sensor = str(filas[0][1]).strip()

        encabezado_unidad = "Value"
        if len(filas) > 1 and len(filas[1]) >= 3:
            encabezado_unidad = str(filas[1][2]).strip()

        datos = filas[2:]

        if promediar:
            grupos = {}
            for fila in datos:
                if len(fila) < 3:
                    continue

                fecha = fila[0]
                hora_simple = self._extraer_hora_simplificada(fila[1])
                valor = fila[2] if isinstance(fila[2], (int, float)) else None

                clave = (fecha, hora_simple)

                if clave not in grupos:
                    grupos[clave] = []

                if valor is not None:
                    grupos[clave].append(valor)

            filas_resultado = []
            for (fecha, hora_simple), valores in grupos.items():
                if valores:
                    prom_val = sum(valores) / len(valores)
                    # Redondeamos a máximo 4 decimales como número real
                    promedio = round(prom_val, 4)
                    
                    # Opcional: Si el número resulta ser entero (ej. 934.0), lo convertimos a int (934)
                    if promedio == int(promedio):
                        promedio = int(promedio)
                else:
                    promedio = ""
                filas_resultado.append((fecha, hora_simple, promedio))
            return nombre_sensor, encabezado_unidad, filas_resultado

        else:
            filas_resultado = []
            for fila in datos:
                if len(fila) >= 3:
                    fecha = fila[0]
                    hora_exacta = str(fila[1]).strip()
                    valor = fila[2]
                    filas_resultado.append((fecha, hora_exacta, valor))
            return nombre_sensor, encabezado_unidad, filas_resultado

    def _obtener_filas_unidas(self):
        """Une los archivos compartiendo las primeras 2 columnas (Date y Time)."""
        if not self.cleaned_files_data:
            return []

        promediar = self.var_promediar.get()
        sensores = []
        unidades = []
        datos_archivos = []

        for f in self.cleaned_files_data:
            sensor, unidad, datos = self._procesar_archivo_individual(f, promediar=promediar)
            sensores.append(sensor)
            unidades.append(unidad)
            datos_archivos.append(datos)

        header_row1 = ["", "Sensor Type"] + sensores
        header_row2 = ["Date", "Time"] + unidades

        parsed_records = []
        for idx_archivo, datos in enumerate(datos_archivos):
            for fecha, hora, valor in datos:
                dt = self._parse_dt(fecha, hora)
                if dt:
                    parsed_records.append((dt, fecha, hora, idx_archivo, valor))

        if not parsed_records:
            return [header_row1, header_row2]

        all_dts = [r[0] for r in parsed_records]
        min_dt = min(all_dts)
        max_dt = max(all_dts)

        fechas_horas_ordenadas = []
        tabla_map = {}

        if promediar:
            curr_dt = datetime(min_dt.year, min_dt.month, min_dt.day, min_dt.hour, 0, 0)
            end_dt = datetime(max_dt.year, max_dt.month, max_dt.day, max_dt.hour, 0, 0)

            while curr_dt <= end_dt:
                fecha_fmt = curr_dt.strftime("%d/%m/%Y")
                hora_fmt = str(curr_dt.hour)
                clave = (fecha_fmt, hora_fmt)
                if clave not in tabla_map:
                    tabla_map[clave] = [""] * len(datos_archivos)
                    fechas_horas_ordenadas.append(clave)
                curr_dt += timedelta(hours=1)

            for dt, fecha, hora, idx_archivo, valor in parsed_records:
                fecha_fmt = dt.strftime("%d/%m/%Y")
                hora_fmt = str(dt.hour)
                clave = (fecha_fmt, hora_fmt)
                if clave in tabla_map:
                    tabla_map[clave][idx_archivo] = valor

        else:
            sorted_records = sorted(parsed_records, key=lambda x: x[0])
            for dt, fecha, hora, idx_archivo, valor in sorted_records:
                clave = (fecha, hora)
                if clave not in tabla_map:
                    tabla_map[clave] = [""] * len(datos_archivos)
                    fechas_horas_ordenadas.append(clave)
                tabla_map[clave][idx_archivo] = valor

        filas_finales = [header_row1, header_row2]
        for fecha, hora in fechas_horas_ordenadas:
            valores_sensores = tabla_map[(fecha, hora)]
            filas_finales.append([fecha, hora] + valores_sensores)

        return filas_finales

    def _generar_nombre_unico(self, extension):
        """Genera un nombre dinámico según la fecha, la hora y si los datos fueron promediados."""
        folder = self.output_folder if self.output_folder and os.path.exists(self.output_folder) else os.getcwd()

        fecha_hora_actual = datetime.now().strftime("%d-%m-%Y_%H-%M")
        
        if self.var_promediar.get():
            base_name = f"union_de_datos_PROMEDIADOS_{fecha_hora_actual}"
        else:
            base_name = f"union_de_datos_{fecha_hora_actual}"

        filename = f"{base_name}{extension}"
        full_path = os.path.join(folder, filename)

        counter = 1
        while os.path.exists(full_path):
            filename = f"{base_name}({counter}){extension}"
            full_path = os.path.join(folder, filename)
            counter += 1

        return folder, filename

    def save_csv(self):
        folder, initial_file = self._generar_nombre_unico(".csv")

        # --- CORRECCIÓN DE LA CARPETA DE GUARDADO ---
        # Pasamos la ruta completa en initialfile para forzar a Windows a abrir 'folder'
        full_initial_path = os.path.join(folder, initial_file)

        save_path = filedialog.asksaveasfilename(
            title="Guardar archivo unido como CSV",
            defaultextension=".csv",
            initialdir=folder,
            initialfile=full_initial_path,  # <--- Ruta completa corregida
            filetypes=[("Archivos CSV", "*.csv")],
        )
        if not save_path:
            return

        try:
            filas_unidas = self._obtener_filas_unidas()

            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=",") # <-- Cambio a coma (,) para correcta separación
                for fila in filas_unidas:
                    writer.writerow(fila)

            self.lbl_status.configure(text="¡CSV guardado con éxito!", text_color="green")
            messagebox.showinfo("Éxito", f"Archivo guardado en:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo CSV:\n{e}")

    def save_excel(self):
        folder, initial_file = self._generar_nombre_unico(".xlsx")

        # --- CORRECCIÓN DE LA CARPETA DE GUARDADO ---
        full_initial_path = os.path.join(folder, initial_file)

        save_path = filedialog.asksaveasfilename(
            title="Guardar archivo unido como Excel",
            defaultextension=".xlsx",
            initialdir=folder,
            initialfile=full_initial_path,  # <--- Ruta completa corregida
            filetypes=[("Libro de Excel", "*.xlsx")],
        )
        if not save_path:
            return

        try:
            filas_unidas = self._obtener_filas_unidas()

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Datos Unidos"

            for fila in filas_unidas:
                ws.append(fila)

            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            wb.save(save_path)
            self.lbl_status.configure(text="¡Excel guardado con éxito!", text_color="green")
            messagebox.showinfo("Éxito", f"Archivo Excel guardado en:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo Excel:\n{e}")


if __name__ == "__main__":
    root = ctk.CTk()
    app = CSVFixerAndMergerApp(root)
    root.mainloop()