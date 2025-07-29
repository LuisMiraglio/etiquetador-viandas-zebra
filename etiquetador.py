# =============================================================================
# IMPORTS Y CONFIGURACIÓN INICIAL
# =============================================================================
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime
import os
import socket
import win32print
import subprocess
import serial  
import json

class EtiquetadoraApp:
    def aplicar_efecto_hover(self, boton, color_normal, color_hover):
        """Aplica efecto de cambio de color al pasar el mouse sobre un botón"""
        
        # Guardar el color original para poder restaurarlo
        boton._color_normal = color_normal
        
        # Función para cuando el mouse entra al botón
        def on_enter(e):
            if boton['state'] != tk.DISABLED:
                boton.config(bg=color_hover)
                
        # Función para cuando el mouse sale del botón
        def on_leave(e):
            if boton['state'] != tk.DISABLED:
                boton.config(bg=color_normal)
                
        # Vincular eventos de mouse
        boton.bind("<Enter>", on_enter)
        boton.bind("<Leave>", on_leave)
        
        # Sobrescribir config para manejar el estado deshabilitado
        original_config = boton.config
        
        def config_override(**kwargs):
            if 'state' in kwargs:
                if kwargs['state'] == tk.DISABLED:
                    original_config(bg="#cccccc", fg="#999999")  # Color gris para deshabilitado
                elif kwargs['state'] == tk.NORMAL:
                    original_config(bg=color_normal, fg="white")  # Restaurar colores originales
            
            # Llamar a la configuración original con todos los argumentos
            return original_config(**kwargs)
            
        # Reemplazar el método config original
        boton.config = config_override
        boton.configure = config_override
    
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Etiqueta")
        self.excel_path = ""
        
        # Inicializar el tipo de conexión
        self.tipo_conexion = tk.StringVar()
        self.tipo_conexion.set("USB")  # Por defecto USB
        
        # Configuración de rutas
        self.config_folder = os.path.join(os.path.expanduser("~"), ".etiquetador")
        self.config_file = os.path.join(self.config_folder, "config.json")
        
        # Crear carpeta de configuración si no existe
        if not os.path.exists(self.config_folder):
            os.makedirs(self.config_folder, exist_ok=True)
        
        # Cargar configuraciones guardadas
        self.cargar_configuraciones()
        
        # Definir constantes de espaciado para consistencia
        PADDING_EXTERNO = 30    # Padding externo para marcos principales
        PADDING_INTERNO = 20    # Padding interno para secciones
        SPACING_VERTICAL = 20   # Espaciado vertical entre secciones
        SPACING_ELEMENTOS = 12  # Espaciado entre elementos dentro de secciones
        WIDGET_PADDING = 8      # Padding estándar para widgets individuales
        
        # Establecer tema de colores y fuentes
        bg_color = "#f9fafc"           # Fondo general blanco muy claro
        header_color = "#003366"       # Azul corporativo oscuro
        section_color = "#f4f6fb"      # Azul/gris muy claro para secciones (más sutil)
        border_color = "#b2becd"       # Bordes suaves
        text_color = "#22223b"         # Texto principal
        accent_color = "#d32f2f"       # Rojo corporativo para botones
        font_main = ("Segoe UI", 11)
        font_title = ("Segoe UI Semibold", 16, "bold")
        font_section = ("Segoe UI", 12, "bold")
        font_label = ("Segoe UI", 10)
        font_button = ("Segoe UI Semibold", 10, "bold")
        
        # Configurar el root para que se expanda
        self.root.columnconfigure(0, weight=1)
        self.root.configure(bg=bg_color)
        
        # Aumentar el espaciado vertical del encabezado para dar más "aire"
        header_frame = tk.Frame(root, bg=header_color)
        header_frame.pack(fill=tk.X)
        title_label = tk.Label(header_frame, 
                              text="SISTEMA DE ETIQUETADO", 
                              font=font_title, 
                              fg="white", 
                              bg=header_color, 
                              pady=24)  # Aumentado para más espacio vertical
        title_label.pack(fill=tk.X)
        
        # Contenedor principal con más separación uniforme
        main_frame = tk.Frame(root, bd=0, relief=tk.FLAT, bg=bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=PADDING_EXTERNO, pady=PADDING_EXTERNO)
        main_frame.columnconfigure(0, weight=1)
        
        # Aplicar espaciado consistente a todas las secciones
        for i in range(4):  # Preparamos 4 filas para las secciones
            main_frame.rowconfigure(i, weight=0, minsize=100)  # Altura mínima para cada sección
        
        # Sección de selección de archivo con espaciado estandarizado
        seccion_archivo = tk.LabelFrame(main_frame, text="Archivo Excel", 
                                      font=font_section, 
                                      padx=PADDING_INTERNO, 
                                      pady=PADDING_INTERNO, 
                                      bg=section_color, 
                                      fg=header_color,
                                      bd=2, relief=tk.GROOVE, 
                                      labelanchor="nw")
        seccion_archivo.grid(row=0, column=0, sticky="ew", pady=(0, SPACING_VERTICAL))
        seccion_archivo.columnconfigure(0, weight=1)
        
        # Frame para botones de Excel
        self.excel_frame = tk.Frame(seccion_archivo, bg=section_color)
        self.excel_frame.grid(row=0, column=0, sticky="ew", pady=5)
        self.excel_frame.columnconfigure(0, weight=1)
        self.excel_frame.columnconfigure(1, weight=1)

        # Botón para seleccionar Excel (sin ícono)
        self.btn_excel = tk.Button(self.excel_frame, 
                                  text="Seleccionar Excel", 
                                  command=self.seleccionar_excel,
                                  bg="#4CAF50",  
                                  fg="white",    
                                  font=("Arial", 10, "bold"),
                                  padx=10,
                                  relief=tk.RAISED,
                                  borderwidth=2,
                                  activebackground="#2e7d32",  # Color mucho más oscuro al hacer clic
                                  activeforeground="white")
        self.btn_excel.grid(row=0, column=0, padx=5, sticky="ew")
        self.aplicar_efecto_hover(self.btn_excel, "#4CAF50", "#2e7d32")  # Verde oscuro para hover
        
        # Botón para eliminar selección de Excel (sin ícono)
        self.btn_eliminar_excel = tk.Button(self.excel_frame, 
                                          text="Eliminar selección", 
                                          command=self.eliminar_seleccion_excel,
                                          bg="#F44336", 
                                          fg="white",    
                                          font=("Arial", 10, "bold"),
                                          padx=10,
                                          relief=tk.RAISED,
                                          borderwidth=2,
                                          activebackground="#b71c1c",  # Rojo mucho más oscuro
                                          activeforeground="white")
        self.btn_eliminar_excel.grid(row=0, column=1, padx=5, sticky="ew")
        self.aplicar_efecto_hover(self.btn_eliminar_excel, "#F44336", "#b71c1c")
        self.btn_eliminar_excel.config(state=tk.DISABLED)
        
        # Etiqueta para mostrar el archivo seleccionado
        self.label_excel_seleccionado = tk.Label(seccion_archivo, 
                                               text="Ningún archivo seleccionado", 
                                               fg="#6c757d", bg=section_color,
                                               font=font_label)
        self.label_excel_seleccionado.grid(row=1, column=0, pady=5, sticky="ew")

        # Sección de fecha - visualmente separada
        seccion_fecha = tk.LabelFrame(main_frame, text="Fecha de vencimiento", 
                                    font=font_section, 
                                    padx=20, pady=15, 
                                    bg=section_color, fg=header_color,
                                    bd=2, relief=tk.GROOVE, labelanchor="nw")
        seccion_fecha.grid(row=1, column=0, sticky="ew", pady=(0, SPACING_VERTICAL))
        seccion_fecha.columnconfigure(0, weight=1)
        
        
        # Frame para el selector de fecha con un mejor estilo
        fecha_entry_frame = tk.Frame(seccion_fecha, bg=bg_color)
        fecha_entry_frame.grid(row=1, column=0, sticky="w", pady=(0, 2))

        # Obtener la fecha actual
        hoy = datetime.now().date()
        
        # Calendario mejorado con estilos personalizados y validación reforzada
        self.fecha_entry = DateEntry(fecha_entry_frame, 
                                   width=12, 
                                   background='#1976D2',
                                   foreground='white', 
                                   borderwidth=2,
                                   date_pattern='dd/mm/yyyy',
                                   font=("Arial", 10, "bold"),
                                   selectbackground='#0D47A1',
                                   selectforeground='white',
                                   normalbackground='#E3F2FD',
                                   normalforeground='#0D47A1',
                                   headersbackground='#1976D2',
                                   headersforeground='white',
                                   weekendbackground='#BBDEFB',
                                   weekendforeground='#0D47A1',
                                   othermonthbackground='#F5F5F5',
                                   othermonthforeground='#9E9E9E',
                                   cursor="hand2",
                                   mindate=hoy)  # Restringir a fecha actual o posterior
        
        # Establecer explícitamente la fecha inicial como hoy
        self.fecha_entry.set_date(hoy)
        self.fecha_entry.pack(side=tk.LEFT)
        
        # Agregar validación después de cambios de fecha
        self.fecha_entry.bind("<<DateEntrySelected>>", self.validar_fecha)
        
        # Se elimina la etiqueta del calendario con ícono
        # Nueva sección específica para selección de impresora
        seccion_impresora = tk.LabelFrame(main_frame, text="Selección de impresora", 
                                        font=font_section, 
                                        padx=20, pady=15, 
                                        bg=section_color, fg=header_color,
                                        bd=2, relief=tk.GROOVE, labelanchor="nw")
        seccion_impresora.grid(row=2, column=0, sticky="ew", pady=(0, SPACING_VERTICAL))
        seccion_impresora.columnconfigure(0, weight=1)
        
        # Frame para selección de impresora (movido a su propia sección)
        impresora_frame = tk.Frame(seccion_impresora, bg=bg_color)
        impresora_frame.pack(fill=tk.X, expand=True, pady=5)
        
        # Etiqueta para impresora
        impresora_label = tk.Label(impresora_frame, 
                                 text="Impresora Zebra:", 
                                 bg=bg_color, fg=text_color)
        impresora_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # Obtener lista de impresoras disponibles
        self.impresoras = self.obtener_impresoras_zebra()
        if not self.impresoras:
            self.impresoras = ["No se encontraron impresoras Zebra"]
            
        # Combobox para selección de impresora
        self.impresora_combo = ttk.Combobox(impresora_frame, 
                                          values=self.impresoras, 
                                          width=25,
                                          state="readonly")
        if self.impresoras:
            self.impresora_combo.current(0)
        self.impresora_combo.pack(side=tk.LEFT)
        
        # Variables para mantener la funcionalidad de selección de puertos COM aunque no sea visible
        self.combobox_com = ttk.Combobox(impresora_frame, state="readonly", width=10)
        self.combobox_com['values'] = self.obtener_puertos_com()
        self.combobox_com.set("COM1")
        # No empaquetamos el combobox_com para mantenerlo oculto

        # Sección de acción - Solo con un botón de imprimir
        seccion_accion = tk.LabelFrame(main_frame, text="Generar etiquetas", 
                                     font=font_section, 
                                     padx=20, pady=15, 
                                     bg=section_color, fg=header_color,
                                     bd=2, relief=tk.GROOVE, labelanchor="nw")
        seccion_accion.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        seccion_accion.columnconfigure(0, weight=1)
        
        # Botón único de imprimir que ocupa todo el ancho
        self.btn_imprimir = tk.Button(seccion_accion, 
                                    text="IMPRIMIR ETIQUETAS", 
                                    command=self.imprimir_directamente,
                                    bg="#007BFF", 
                                    fg="white",    
                                    font=("Arial", 12, "bold"),
                                    padx=30, pady=15,
                                    relief=tk.RAISED,
                                    borderwidth=2,
                                    activebackground="#0056b3",
                                    activeforeground="white")
        self.btn_imprimir.pack(fill=tk.X, expand=True, padx=10, pady=15)
        self.aplicar_efecto_hover(self.btn_imprimir, "#007BFF", "#0056b3")
        
        # Pie de página con créditos
        footer_frame = tk.Frame(root, bg=header_color)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        credit_label = tk.Label(footer_frame,                        
                              fg="white", bg=header_color, font=("Segoe UI", 9, "italic"), pady=6)
        credit_label.pack(side=tk.RIGHT, padx=10)
        
        # Configurar un tamaño mínimo para la ventana
        self.root.minsize(500, 450)
        
        # Si la ventana se redimensiona, centrar los elementos
        for i in range(4):  # Ahora tenemos 4 filas en lugar de 3
            main_frame.rowconfigure(i, weight=1)
        
        # Centrar la ventana en la pantalla
        self.center_window()
    
    # =============================================================================
    # MÉTODOS DE INTERFAZ Y UTILIDADES
    # =============================================================================
    
    def center_window(self):
        """Centra la ventana en la pantalla."""
        # Centrar la ventana en la pantalla
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
    def validar_fecha(self, event=None):
        """Valida que la fecha seleccionada no sea anterior a hoy."""
        fecha_seleccionada = self.fecha_entry.get_date()
        hoy = datetime.now().date()
        
        if fecha_seleccionada < hoy:
            messagebox.showerror("Error de fecha", "No se puede seleccionar una fecha anterior a hoy.")
            # Establecer la fecha de nuevo a hoy
            self.fecha_entry.set_date(hoy)
            return
            
    # =============================================================================
    # MÉTODOS DE MANEJO DE ARCHIVOS EXCEL
    # =============================================================================
    
    def seleccionar_excel(self):
        """Abre un diálogo para seleccionar un archivo Excel."""
        # Usar la última carpeta como inicio
        inicio_dir = self.configuraciones.get("ultima_carpeta_excel", os.path.expanduser("~"))
        
        nuevo_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            initialdir=inicio_dir
        )
        
        if nuevo_path:
            self.excel_path = nuevo_path
            # Mostrar solo el nombre del archivo, no la ruta completa
            nombre_archivo = os.path.basename(self.excel_path)
            self.label_excel_seleccionado.config(text=f"Archivo: {nombre_archivo}", fg="black")
            self.btn_eliminar_excel.config(state=tk.NORMAL)  # Habilitar botón de eliminación
            # MENSAJE: Archivo cargado
            messagebox.showinfo("Archivo cargado", f"Seleccionado: {nombre_archivo}")

    def eliminar_seleccion_excel(self):
        """Elimina la selección actual del archivo Excel."""
        self.excel_path = ""
        self.label_excel_seleccionado.config(text="Ningún archivo seleccionado", fg="gray")
        self.btn_eliminar_excel.config(state=tk.DISABLED)  # Deshabilitar botón de eliminación
        # MENSAJE: Selección eliminada
        messagebox.showinfo("Selección eliminada", "Se ha eliminado la selección del archivo Excel")

    # =============================================================================
    # MÉTODOS DE CONEXIÓN CON IMPRESORAS
    # =============================================================================
    
    def on_tipo_conexion_change(self, event=None):
        """Maneja el cambio del tipo de conexión de la impresora."""
        if self.tipo_conexion.get() == "Serie":
            self.combobox_com['values'] = self.obtener_puertos_com()
        # No hacemos pack ni unpack ya que la interfaz nunca se muestra

    # El método verificar_conexion_impresora se modifica para uso interno sin interfaz
    def verificar_conexion_impresora(self, event=None, show_always=False):
        """Verifica si la impresora Zebra está correctamente conectada (uso interno)."""
        tipo = self.tipo_conexion.get()
        impresora = self.impresora_combo.get()
        
        if impresora == "No se encontraron impresoras Zebra":
            return False
            
        try:
            if tipo == "Serie":
                # Para conexión serie, intentamos abrir el puerto
                puerto = self.combobox_com.get()
                with serial.Serial(port=puerto, baudrate=9600, timeout=1) as ser:
                    pass
                return True
            else:
                # Para USB/Paralelo verificamos si podemos abrir la impresora
                handle = win32print.OpenPrinter(impresora)
                win32print.ClosePrinter(handle)
                return True
        except Exception as e:
            if show_always:
                messagebox.showerror("Error de conexión", 
                                  f"No se puede conectar con la impresora.\nVerifique que esté encendida y conectada correctamente.")
            return False
        
    def obtener_puertos_com(self):
        """Devuelve la lista de puertos COM disponibles en el sistema."""
        import serial.tools.list_ports
        return [port.device for port in serial.tools.list_ports.comports()]

    def obtener_impresoras_zebra(self):
        """Obtiene la lista de impresoras Zebra instaladas en el sistema."""
        impresoras_zebra = []
        try:
            # Buscar impresoras locales (USB y Paralelo)
            for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL):
                printer_name = printer[2].lower()
                if any(term in printer_name for term in ["zebra", "zdesigner", "zt", "gk", "zd", "lp", "gx", "gc"]):
                    impresoras_zebra.append(printer[2])
        except Exception as e:
            print(f"Error al enumerar impresoras: {e}")
            try:
                output = subprocess.check_output("wmic printer get name", shell=True).decode('utf-8')
                for line in output.split('\n'):
                    line = line.strip().lower()
                    if any(term in line for term in ["zebra", "zdesigner", "zt", "gk", "zd"]):
                        impresoras_zebra.append(line)
            except:
                pass
        return impresoras_zebra if impresoras_zebra else ["No se encontraron impresoras Zebra"]
        
    def refrescar_impresoras(self):
        """Actualiza la lista de impresoras disponibles."""
        impresora_actual = self.impresora_combo.get()
        self.impresoras = self.obtener_impresoras_zebra()
        self.impresora_combo['values'] = self.impresoras
        
        # Intentar mantener la impresora seleccionada si todavía existe
        if impresora_actual in self.impresoras:
            self.impresora_combo.set(impresora_actual)
        elif self.impresoras:
            self.impresora_combo.current(0)
    
    # =============================================================================
    # MÉTODOS DE IMPRESIÓN
    # =============================================================================
    
    def imprimir_directamente(self):
        """Imprime directamente en la impresora Zebra seleccionada."""
        # Cambiar cursor a espera y desactivar botón para evitar doble click
        self.root.config(cursor="watch") 
        self.btn_imprimir.config(state=tk.DISABLED)
        self.root.update()
        
        try:
            # Serie de validaciones previas a la impresión
            if not self.excel_path:
                # MENSAJE: Error - Selecciona un archivo Excel primero
                messagebox.showerror("Error", "Selecciona un archivo Excel primero.")
                return
                
            impresora = self.impresora_combo.get()
            if impresora == "No se encontraron impresoras Zebra" or not impresora:
                # MENSAJE: Error - No hay impresora Zebra seleccionada
                messagebox.showerror("Error", "No hay impresora Zebra seleccionada.")
                return
                
            # Verificar si la impresora seleccionada sigue disponible
            impresoras_actuales = self.obtener_impresoras_zebra()
            if impresora not in impresoras_actuales:
                # MENSAJE: Pregunta - Impresora no disponible
                respuesta = messagebox.askquestion("Impresora no disponible", 
                                                f"La impresora '{impresora}' ya no está disponible.\n¿Desea actualizar la lista de impresoras?")
                if respuesta == 'yes':
                    self.refrescar_impresoras()
                return
                
            # Validar que el archivo Excel siga existiendo
            if not os.path.exists(self.excel_path):
                # MENSAJE: Error - El archivo Excel seleccionado ya no existe
                messagebox.showerror("Error", "El archivo Excel seleccionado ya no existe. Por favor, seleccione otro archivo.")
                self.excel_path = ""
                self.label_excel_seleccionado.config(text="Ningún archivo seleccionado", fg="gray")
                self.btn_eliminar_excel.config(state=tk.DISABLED)
                return
                
            try:
                df = pd.read_excel(self.excel_path, header=1)
            except Exception as e:
                # MENSAJE: Error - No se pudo leer el Excel
                messagebox.showerror("Error", f"No se pudo leer el Excel: {e}")
                return
        
            if df.empty:
                # MENSAJE: Error - El archivo Excel está vacío
                messagebox.showerror("Error", "El archivo Excel está vacío.")
                return
                
            # Verificar si hay al menos una fila con código válido
            codigos_validos = 0
            for _, row in df.iterrows():
                codigo = str(row.get("Código del menú", "")).strip()
                if codigo:
                    codigos_validos += 1
                    
            if codigos_validos == 0:
                # MENSAJE: Error - El Excel no contiene códigos de menú válidos
                messagebox.showerror("Error", "El Excel no contiene códigos de menú válidos.")
                return

            # A partir de aquí, continúa el procesamiento original
            fecha_vencimiento = self.fecha_entry.get_date()
            fecha_elaboracion = datetime.now()
            
            # Procesar todas las filas del Excel
            etiquetas_generadas = 0
            etiquetas_enviadas = 0
            
            # Mostrar una ventana de progreso
            progreso = tk.Toplevel(self.root)
            progreso.title("Imprimiendo etiquetas")
            progreso.transient(self.root)
            progreso.grab_set()
            
            # Centrar ventana de progreso
            progreso_width = 300
            progreso_height = 100
            progreso_x = self.root.winfo_x() + (self.root.winfo_width() - progreso_width) // 2
            progreso_y = self.root.winfo_y() + (self.root.winfo_height() - progreso_height) // 2
            progreso.geometry(f"{progreso_width}x{progreso_height}+{progreso_x}+{progreso_y}")
            
            progreso.configure(bg="#f0f0f0")
            progreso.resizable(False, False)
            
            mensaje_label = tk.Label(progreso, text=f"Imprimiendo etiquetas: 0/{codigos_validos}", 
                                  font=("Arial", 10), bg="#f0f0f0", pady=10)
            mensaje_label.pack()
            
            barra_progreso = ttk.Progressbar(progreso, orient="horizontal", 
                                          length=250, mode="determinate", maximum=codigos_validos)
            barra_progreso.pack(pady=5)
            
            # Actualizar la interfaz para mostrar la ventana de progreso
            self.root.update()
            
            for index, row in df.iterrows():
                codigo = str(row.get("Código del menú", "")).strip()
                if not codigo:  # Omitir filas sin código
                    continue
                    
                # Rellenar con ceros a la izquierda hasta tener 12 dígitos para EAN-13
                # EAN-13 necesita exactamente 12 dígitos + 1 dígito de verificación automático
                codigo = codigo.zfill(12)
                
                # Formatear el código EAN-13 con espacios específicos: "X XXXXXX XXXXXX"
                # Pero ahora usamos los 12 dígitos completos
                codigo_formateado = f"{codigo[0]} {codigo[1:7]} {codigo[7:12]}"
                
                # Manejar nombre del menú largo dividiéndolo en dos líneas si es necesario
                nombre_menu_original = str(row.get("Nombre del menú", "")).strip()
                if not nombre_menu_original:
                    nombre_menu_original = "Sin especificar"
                    
                max_caracteres_linea = 30  # Máximo de caracteres por línea
                
                if len(nombre_menu_original) <= max_caracteres_linea:
                    # Si el nombre cabe en una línea, lo usamos tal cual
                    nombre_menu_linea1 = nombre_menu_original
                    nombre_menu_linea2 = ""
                else:
                    # Mejorado: Buscar un punto óptimo para dividir el texto
                    # Inicializar mitad con un valor por defecto seguro
                    mitad = min(max_caracteres_linea, len(nombre_menu_original)-1)
                    
                    # Buscar un espacio cerca del punto de corte deseado
                    # Solo buscar si hay suficiente texto
                    if len(nombre_menu_original) > 1:
                        # Buscar el último espacio antes del límite
                        while mitad > 0 and nombre_menu_original[mitad] != ' ':
                            mitad -= 1
                        
                        # Si no hay espacios cerca, simplemente cortar en el límite
                        if mitad <= 5:  # Si está muy al principio, mejor cortar en el límite
                            mitad = min(max_caracteres_linea, len(nombre_menu_original)-1)
                    
                    nombre_menu_linea1 = nombre_menu_original[:mitad].strip()
                    nombre_menu_linea2 = nombre_menu_original[mitad:].strip()
                    
                    # Si la segunda línea es muy larga, la truncamos
                    if len(nombre_menu_linea2) > max_caracteres_linea:
                        nombre_menu_linea2 = nombre_menu_linea2[:max_caracteres_linea] + "..."
                
                nombre_empleado = str(row.get("Nombre de empleado", "")).strip()
                if not nombre_empleado:
                    nombre_empleado = "Sin especificar"

                # Convertir el nombre a formato "APELLIDO, NOMBRE" en mayúsculas
                # Si el nombre ya tiene coma, usar tal como está, sino convertir
                if ',' in nombre_empleado:
                    nombre_empleado_formato = nombre_empleado.upper()
                else:
                    # Intentar separar nombre y apellido para formato "APELLIDO, NOMBRE"
                    partes_nombre = nombre_empleado.strip().split()
                    if len(partes_nombre) >= 2:
                        # Asumir que la última palabra es el apellido
                        apellido = partes_nombre[-1]
                        nombres = ' '.join(partes_nombre[:-1])
                        nombre_empleado_formato = f"{apellido.upper()}, {nombres.upper()}"
                    else:
                        nombre_empleado_formato = nombre_empleado.upper()

                # Generar ZPL con el formato exacto de la etiqueta real
                    zpl = f"""
^XA
^PW400
^LH20,20
^CI28

^CF0,30
^FO0,5^FB360,40,1,C^FDLUGAR: Comedor Bella Italia^FS
^FO0,35^GB360,2,2^FS  ; Línea horizontal como subrayado

^CF0,25
^FO10,45^FD{nombre_empleado_formato}^FS
^FO10,80^FDMenu: {nombre_menu_linea1}^FS
^FO10,115^FDELAB: {fecha_elaboracion.strftime('%d/%m/%Y')}^FS
^FO10,150^FDVENC: {fecha_vencimiento.strftime('%d/%m/%Y')}^FS

^BY2,3.0,120
^FO40,175^BEN,120,Y,N^FD{codigo}^FS

^XZ
"""



                # Intentar enviar a la impresora directamente
                try:
                    tipo = self.tipo_conexion.get()
                    if tipo == "USB" or tipo == "Paralelo":
                        self.enviar_a_impresora(zpl, impresora)
                    elif tipo == "Serie":
                        puerto = self.combobox_com.get()
                        self.enviar_por_serie(zpl, puerto)
                    else:
                        raise Exception("Tipo de conexión no soportado.")
                    etiquetas_enviadas += 1
                    # Actualizar barra de progreso
                    barra_progreso["value"] = etiquetas_enviadas
                    mensaje_label.config(text=f"Imprimiendo etiquetas: {etiquetas_enviadas}/{codigos_validos}")
                    self.root.update()
                except Exception as e:
                    # MENSAJE: Error de impresión - No se pudo imprimir la etiqueta para X
                    messagebox.showerror("Error de impresión", f"No se pudo imprimir la etiqueta para {nombre_empleado}: {str(e)}")
                etiquetas_generadas += 1

            # MENSAJE: Impresión completada o Aviso
            if etiquetas_enviadas > 0:
                messagebox.showinfo("Impresión completada", 
                                  f"Se imprimieron {etiquetas_enviadas} de {etiquetas_generadas} etiquetas en la impresora '{impresora}'.")
            else:
                messagebox.showwarning("Aviso", "No se pudieron imprimir etiquetas en la impresora.")

        except Exception as e:
            # MENSAJE: Error inesperado
            messagebox.showerror("Error inesperado", f"Ocurrió un error durante el proceso: {str(e)}")
        finally:
            # Cerrar ventana de progreso si sigue abierta
            try:
                if 'progreso' in locals() and progreso.winfo_exists():
                    barra_progreso.stop()
                    progreso.destroy()
            except:
                pass
                
            # Restaurar interfaz
            self.root.config(cursor="")
            self.btn_imprimir.config(state=tk.NORMAL)
            self.root.update()
    
    def enviar_a_impresora(self, zpl_data, impresora):
        """Envía el código ZPL a la impresora Zebra por USB o puerto paralelo."""
        handle = None
        try:
            handle = win32print.OpenPrinter(impresora)
            job = win32print.StartDocPrinter(handle, 1, ("Etiqueta ZPL", None, "RAW"))
            win32print.StartPagePrinter(handle)
            win32print.WritePrinter(handle, zpl_data.encode())
            win32print.EndPagePrinter(handle)
            win32print.EndDocPrinter(handle)
            win32print.ClosePrinter(handle)
            handle = None
            return True
        except Exception as e:
            if handle:
                try:
                    win32print.ClosePrinter(handle)
                except:
                    pass
            error_msg = str(e)
            raise Exception(f"Error al imprimir: Verifica que la impresora esté conectada y encendida. {error_msg}")

    def enviar_por_serie(self, zpl_data, puerto):
        """Envía el código ZPL por puerto serie (RS-232)."""
        try:
            # Configuración típica: 9600 baudios, 8N1, sin control de flujo
            with serial.Serial(port=puerto, baudrate=9600, bytesize=8, parity='N', stopbits=1, timeout=2) as ser:
                ser.write(zpl_data.encode('utf-8'))
        except Exception as e:
            raise Exception(f"Error al enviar por puerto serie {puerto}: {e}")
    
    # =============================================================================
    # MÉTODOS DE CARGA Y GUARDADO DE CONFIGURACIONES
    # =============================================================================
    
    def cargar_configuraciones(self):
        """Carga las configuraciones guardadas."""
        self.configuraciones = {
            "ultima_impresora": "",
            "ultimo_tipo_conexion": "USB",
            "ultimo_puerto_com": "COM1",
            "ultima_carpeta_excel": os.path.expanduser("~"),
            "recientes": []
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    saved_config = json.load(f)
                    # Actualizar configuraciones con las guardadas
                    self.configuraciones.update(saved_config)
        except Exception as e:
            print(f"Error al cargar configuraciones: {e}")
    
    def guardar_configuraciones(self):
        """Guarda las configuraciones actuales."""
        try:
            # Actualizar configuraciones antes de guardar
            if hasattr(self, 'impresora_combo'):
                self.configuraciones["ultima_impresora"] = self.impresora_combo.get()
            
            if hasattr(self, 'tipo_conexion'):
                self.configuraciones["ultimo_tipo_conexion"] = self.tipo_conexion.get()
            
            if hasattr(self, 'combobox_com'):
                self.configuraciones["ultimo_puerto_com"] = self.combobox_com.get()
            
            # Guardar la última carpeta usada para abrir Excel
            if self.excel_path:
                self.configuraciones["ultima_carpeta_excel"] = os.path.dirname(self.excel_path)
                
                # Agregar a recientes si no está ya
                if self.excel_path not in self.configuraciones["recientes"]:
                    self.configuraciones["recientes"].insert(0, self.excel_path)
                    # Mantener solo los últimos 5 archivos
                    self.configuraciones["recientes"] = self.configuraciones["recientes"][:5]
            
            with open(self.config_file, 'w') as f:
                json.dump(self.configuraciones, f)
        except Exception as e:
            print(f"Error al guardar configuraciones: {e}")
    
    # =============================================================================
    # MÉTODOS DE MANIPULACIÓN DE REGISTROS
    # =============================================================================
    
    def setup_treeview(self):
        """Configura el treeview para mostrar los registros del Excel."""
        # Crear scrollbar para la lista
        scrollbar = ttk.Scrollbar(self.registros_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Crear el treeview con columnas
        self.tree = ttk.Treeview(self.registros_frame, 
                               columns=("nombre", "menu"), 
                               show="headings",
                               height=6,
                               selectmode="extended",
                               yscrollcommand=scrollbar.set)
        
        # Configurar columnas
        self.tree.heading("nombre", text="Nombre de empleado")
        self.tree.heading("menu", text="Menú")
        
        self.tree.column("nombre", width=150, anchor="w")
        self.tree.column("menu", width=250, anchor="w")
        
        # Conectar scrollbar
        scrollbar.configure(command=self.tree.yview)
        
        # Empaquetar treeview
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Botones para manipular selección
        btn_frame = tk.Frame(self.registros_frame, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        btn_seleccionar_todo = tk.Button(btn_frame, text="Seleccionar todos", 
                                        command=lambda: self.seleccionar_items(True),
                                        bg="#4CAF50", fg="white")
        btn_seleccionar_todo.pack(side=tk.LEFT, padx=(0, 5))
        
        btn_deseleccionar = tk.Button(btn_frame, text="Deseleccionar todos", 
                                     command=lambda: self.seleccionar_items(False),
                                     bg="#F44336", fg="white")
        btn_deseleccionar.pack(side=tk.LEFT)
        
    def seleccionar_items(self, seleccionar=True):
        """Selecciona o deselecciona todos los items del treeview."""
        if seleccionar:
            self.tree.selection_set(self.tree.get_children())
        else:
            self.tree.selection_remove(self.tree.get_children())
    
    def actualizar_modo_impresion(self):
        """Actualiza la interfaz según el modo de impresión seleccionado."""
        if self.modo_impresion.get() == "seleccion":
            # Cargar los datos del Excel para mostrar en la lista
            self.cargar_datos_excel()
            self.registros_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        else:
            self.registros_frame.pack_forget()
    
    def cargar_datos_excel(self):
        """Carga los datos del excel en el treeview."""
        # Limpiar treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not self.excel_path or not os.path.exists(self.excel_path):
            return
            
        try:
            df = pd.read_excel(self.excel_path, header=1)
            if df.empty:
                return
                
            # Iterar por las filas e insertar en el treeview
            for index, row in df.iterrows():
                codigo = str(row.get("Código del menú", "")).strip()
                if not codigo:  # Omitir filas sin código
                    continue
                    
                # Insertar fila en el treeview
                nombre_empleado = str(row.get("Nombre de empleado", "")).strip() or "Sin especificar"
                nombre_menu = str(row.get("Nombre del menú", "")).strip() or "Sin especificar"
                
                # Truncar textos largos
                if len(nombre_menu) > 40:
                    nombre_menu = nombre_menu[:37] + "..."
                
                self.tree.insert("", tk.END, iid=str(index), values=(nombre_empleado, nombre_menu),
                                tags=(str(index), codigo))
                
            # Seleccionar todos los items por defecto
            self.seleccionar_items(True)
            
        except Exception as e:
            # MENSAJE: Error - No se pudieron cargar los datos del Excel
            messagebox.showerror("Error", f"No se pudieron cargar los datos del Excel: {str(e)}")
    
    # =============================================================================
    # MÉTODOS DE CIERRE Y SALIDA
    # =============================================================================
    
    # Método que se ejecuta al cerrar la aplicación
    def on_close(self):
        """Método llamado al cerrar la aplicación."""
        self.guardar_configuraciones()
        self.root.destroy()

    def cargar_iconos(self):
        """Este método está vacío ya que no usamos iconos externos."""
        pass

# =============================================================================
# PUNTO DE ENTRADA DE LA APLICACIÓN
# =============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    
    # Establecer el icono de la aplicación
    try:
        # Rutas posibles para el icono (en el mismo directorio que el script)
        possible_icons = ["icono.ico", "etiqueta_icon.ico", "app_icon.ico", "zebra.ico", "label.ico"]
        icon_found = False
        
        for icon_name in possible_icons:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), icon_name)
            if os.path.exists(icon_path):
                root.iconbitmap(icon_path)
                icon_found = True
                print(f"Usando icono: {icon_path}")
                break
        
        if not icon_found:
            print("No se encontró ningún archivo de icono en la carpeta del programa.")
            
    except Exception as e:
        print(f"Error al cargar el icono: {e}")
    
    app = EtiquetadoraApp(root)
    # Configurar evento de cierre
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()
    
    
# =============================================================================
# COMANDO PARA COMPILAR EL EJECUTABLE
# =============================================================================
# 1. Activar entorno Python 3.8
#venv_py38\Scripts\activate
#ultimo camando que se uso para compilar
#pyinstaller --onefile --windowed --icon=icono.ico --collect-all babel --collect-all tkcalendar --noupx etiquetador.py