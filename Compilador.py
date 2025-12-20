# CÓDIGO FINAL CORREGIDO Y CON COMPILACIÓN CONDICIONAL
#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import os
import sys
import shutil
import threading
import time
from pathlib import Path
import builtins
import base64
import importlib 

# ====================================================
# IMPORTS CONDICIONALES Y DE SONIDO
# ====================================================

# IMPORTACIÓN REQUERIDA PARA PROCESAMIENTO DE IMÁGENES (Pillow/PIL)
try:
    from PIL import Image
except ImportError:
    Image = None

# Imports para sonido (específico de Windows)
if sys.platform.startswith("win"):
    try:
        import winsound
    except ImportError:
        winsound = None
else:
    winsound = None # Asegura que sea None en Mac/Linux


# ====================================================
# CONFIGURACIÓN DEL COMPILADOR
# ====================================================

# ⚠️ RUTAS DE MACOS: Solo usadas para cross-compiling, ignoradas en Windows.
PYTHON_SISTEMA = "/usr/local/bin/python3"

# Ruta del ejecutable de PyInstaller (ajusta si es necesario)
PYINSTALLER_PATH = "/Library/Frameworks/Python.framework/Versions/3.12/bin/pyinstaller"

# Carpeta donde se guardarán los compilados finales (por defecto)
CARPETA_SALIDA_DEFECTO = os.path.expanduser("~/Desktop/Compilados")

# Paquetes requeridos que serán verificados al inicio
REQUIRED_PACKAGES = ["pyinstaller", "Pillow"] 

# ====================================================
# VARIABLES GLOBALES Y ESTILOS
# ====================================================
# Paleta de colores
COLOR_BG = "#0b2b4a"          # Azul marino oscuro (fondo)
COLOR_FG = "#ffffff"          # Blanco (texto)
COLOR_ACCENT_DARK = "#2aa1d6" # Azul celeste oscuro (acento)
COLOR_ACCENT_LIGHT = "#5ac8fa" # Azul cielo (botón redondeado/texto legible)
COLOR_BOX_BG = "#083047"      # Fondo caja debug/adjuntos
COLOR_BTN_COMPILE = "#1a5276" # Azul semi oscuro para el botón Compilar
COLOR_SUCCESS = "#28a745"     # Verde
COLOR_ERROR = "#dc3545"       # Rojo

# Factor de escala (aproximadamente 30% más grande)
SCALE_FACTOR = 1.3
DEFAULT_FONT_SIZE = int(12 * SCALE_FACTOR)
HEADER_FONT_SIZE = int(16 * SCALE_FACTOR)
BIG_FONT_SIZE = int(13 * SCALE_FACTOR)
BUTTON_FONT_SIZE = int(12 * SCALE_FACTOR)
DEBUG_FONT_SIZE = int(11 * SCALE_FACTOR)
ENTRY_WIDTH = 40 # Ancho de Entry

# Título centrado de la App
APP_TITLE = "Compilador Mac & Win 2025 .: By Morrigan :."

# Lista de archivos temporales creados (para cleanup)
TEMP_ICON_FILES = []

# ====================================================
# UTIL: Funciones de Sonido
# ====================================================

def play_snake_hiss():
    """Reproduce un sonido de 'siseo' (hiss) al finalizar la compilación (Windows)."""
    if sys.platform.startswith("win") and winsound:
        try:
            # Frecuencia alta y corta para simular un hiss (ej. 8000 Hz por 200 ms)
            winsound.Beep(8000, 200) 
        except Exception:
            # Si Beep falla, usar el sonido de sistema por defecto.
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
    elif sys.platform.startswith("darwin"):
        # Sonido simple para Mac
        try:
            os.system('afplay /System/Library/Sounds/Tink.aiff')
        except Exception:
             pass 
    else:
        # Sonido simple de terminal
        print("\a")


# ====================================================
# WIDGET PERSONALIZADO: Botón Redondeado
# ====================================================
class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command=None, radius=20, width=120, height=40,
                 fill=COLOR_ACCENT_LIGHT, fg="black", font=("Helvetica", BUTTON_FONT_SIZE), **kwargs):
        super().__init__(parent, width=width, height=height, bg=parent['bg'], highlightthickness=0, **kwargs)
        self.radius = radius
        self.command = command
        self._initial_fill = fill # Guardar el color inicial
        self._initial_fg = fg      # Guardar el color inicial del texto
        self.fill = fill
        self.text = text
        self.fg = fg # Negro
        self.font = font
        self._state = "normal"

        self.coords = (0, 0, width, height)
        self.draw_rounded_rectangle(self.coords, self.radius)
        
        # 🟢 CORRECCIÓN: Usar self.coords para la posición central
        self.text_id = self.create_text(self.coords[2]/2, self.coords[3]/2, text=text, fill=fg, font=font)
        
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def draw_rounded_rectangle(self, rect, r):
        x1, y1, x2, y2 = rect
        points = [
            x1+r, y1, x2-r, y1,
            x2, y1, x2, y1+r, 
            x2, y2-r, x2, y2, 
            x2-r, y2, x1+r, y2, 
            x1, y2, x1, y2-r,
            x1, y1+r, x1, y1
        ]
        self.delete("all")
        self.create_polygon(points, fill=self.fill, smooth=True)
        # 🟢 CORRECCIÓN: Usar self.coords[2]/2 y self.coords[3]/2 para el centro
        self.text_id = self.create_text(self.coords[2]/2, self.coords[3]/2, text=self.text, fill=self.fg, font=self.font) 

    def _on_click(self, event):
        if self.command and self._state == "normal":
            self.command()

    def _on_enter(self, event):
        if self._state == "normal":
            # Usar COLOR_ACCENT_DARK como hover (definido globalmente)
            self.itemconfig(self.find_all()[0], fill=COLOR_ACCENT_DARK) 

    def _on_leave(self, event):
        if self._state == "normal":
            self.itemconfig(self.find_all()[0], fill=self.fill)

    def set_state(self, state):
        self._state = state
        if state == "disabled":
            self.itemconfig(self.find_all()[0], fill="#888888")
            self.itemconfig(self.text_id, fill="#333333")
        else:
            # Restaurar el color normal (usando self.fill actual)
            self.itemconfig(self.find_all()[0], fill=self.fill)
            self.itemconfig(self.text_id, fill=self.fg)

# ====================================================
# UTIL: Clase para reportar progreso en debug
# ====================================================
class GuiWriter:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.lock = threading.Lock()
        
        # 🟢 CORRECCIÓN VISUAL: Textos de alto contraste sobre fondo oscuro (Debug Principal)
        self.text_widget.tag_configure("progress", foreground=COLOR_ACCENT_LIGHT, font=("Courier", DEBUG_FONT_SIZE, "bold")) 
        self.text_widget.tag_configure("normal", foreground=COLOR_FG, font=("Courier", DEBUG_FONT_SIZE)) # <-- Texto Blanco
        self.progress_percent = 0
        self.total_steps = 100 

    def write(self, s):
        if not s: return
        
        if "building EXE from" in s.lower() or "building app from" in s.lower():
            self.report_progress(10, "[10%] Preparando PyInstaller...")
        elif "checking analysis" in s.lower():
            self.report_progress(20, "[20%] Analizando dependencias...")
        elif "bootloader" in s.lower():
            self.report_progress(40, "[40%] Construyendo bootloader...")
        elif "archiving" in s.lower() or "creating bundle" in s.lower():
            self.report_progress(70, "[70%] Empaquetando archivos y creando bundle...")
        elif "finished" in s.lower():
            self.report_progress(95, "[95%] Finalizando proceso...")
        
        def inner():
            try:
                # Usar el tag "normal" para texto estándar
                self.text_widget.insert(tk.END, s, "normal") 
                self.text_widget.see(tk.END)
            except Exception:
                pass
        try:
            self.text_widget.after(0, inner)
        except Exception:
            pass

    def report_progress(self, percent, message):
        self.progress_percent = percent
        def inner_report():
             self.text_widget.insert(tk.END, f"\n--- {message} ---\n", "progress") 
             self.text_widget.see(tk.END)
        try:
            self.text_widget.after(0, inner_report)
        except Exception:
            pass

    def flush(self):
        pass

# ====================================================
# UTIL: Funciones de limpieza y procesamiento
# ====================================================

def move_file_and_cleanup(carpeta_temp, temp_file_path, extension, output_folder, nombre_base, writer):
    """
    Mueve el archivo final (.app o .exe) y elimina la carpeta temporal.
    """
    try:
        source_dir = os.path.join(carpeta_temp, "dist")
        
        if extension == ".app":
            source_item = os.path.join(source_dir, f"{nombre_base}.app")
            dest_item = os.path.join(output_folder, f"{nombre_base}.app")
            
            if os.path.exists(source_item):
                if os.path.exists(dest_item):
                    shutil.rmtree(dest_item) 
                shutil.move(source_item, dest_item)
                writer.write(f"\n[OK] Aplicación final movida a: {dest_item}\n")
                return 0
            else:
                writer.write("\n[ERROR] No se encontró la carpeta .app final.\n")
                return 1
            
        elif extension == ".exe":
            # 1. Intentar encontrar el archivo con extensión .exe (si se compiló en Windows)
            source_item_exe = os.path.join(source_dir, f"{nombre_base}.exe")
            # 2. Intentar encontrar el archivo sin extensión (común al compilar --onefile en Mac/Linux)
            source_item_noext = os.path.join(source_dir, nombre_base)
            
            final_source_path = None

            if os.path.exists(source_item_exe):
                final_source_path = source_item_exe
            elif os.path.exists(source_item_noext):
                final_source_path = source_item_noext
            
            dest_item = os.path.join(output_folder, f"{nombre_base}.exe")
            
            if final_source_path:
                shutil.move(final_source_path, dest_item)
                writer.write(f"\n[OK] Ejecutable final movido a: {dest_item}\n")
                return 0
            else:
                writer.write(f"\n[ERROR] No se encontró el archivo '{nombre_base}' (o '{nombre_base}.exe') en dist.\n")
                return 1
        return 1
    except Exception as e:
        writer.write(f"\n[ERROR al mover/limpiar]: {e}\n")
        return 1
    finally:
        # LIMPIEZA: Eliminar la carpeta temporal COMPLETA
        try:
            if os.path.exists(carpeta_temp):
                writer.write(f"\n[LIMPIEZA] Eliminando carpeta temporal: {os.path.basename(carpeta_temp)}\n")
                shutil.rmtree(carpeta_temp)
        except Exception as e:
            writer.write(f"\n[ADVERTENCIA] No se pudo limpiar la carpeta temporal: {e}\n")


def process_icon_file(original_path, output_folder):
    """
    Convierte PNG/JPG/ICNS a ICO (para Windows) o ICNS (para Mac) en la carpeta temporal.
    """
    global TEMP_ICON_FILES
    
    if not Image:
        return True 

    path_obj = Path(original_path)
    suffix = path_obj.suffix.lower()
    base_name = path_obj.stem
    temp_folder = os.path.join(output_folder, "temp_icons")
    os.makedirs(temp_folder, exist_ok=True)

    if suffix == '.ico' or suffix == '.icns':
        return True 

    try:
        img = Image.open(original_path)
        
        # Generar ICO para Windows
        temp_ico_path = os.path.join(temp_folder, f"{base_name}.ico")
        img.save(temp_ico_path, format="ICO", sizes=[(256, 256), (48, 48), (32, 32), (16, 16)])
        TEMP_ICON_FILES.append(temp_ico_path)

        # Generar ICNS para Mac
        temp_icns_path = os.path.join(temp_folder, f"{base_name}.icns")
        try:
            icon_sizes = [(16, 16), (32, 32), (64, 64), (128, 128), (256, 256), (512, 512)]
            img.save(temp_icns_path, format="ICNS", sizes=icon_sizes)
            TEMP_ICON_FILES.append(temp_icns_path)
        except Exception:
             pass 
        
        return True
        
    except Exception as e:
        messagebox.showerror("Error de Procesamiento de Ícono", f"No se pudo procesar el archivo {original_path} (Pillow error): {e}")
        return False 

def cleanup_temp_icons():
    """
    Elimina los archivos de ícono temporales generados por Pillow y la carpeta temp_icons.
    """
    global TEMP_ICON_FILES
    for f in TEMP_ICON_FILES:
        try:
            if os.path.exists(f):
                os.remove(f)
        except Exception:
            pass
    TEMP_ICON_FILES = []
    
    # LIMPIEZA DE LA CARPETA TEMP_ICONS SOLICITADA
    temp_folder = os.path.join(CARPETA_SALIDA_DEFECTO, "temp_icons")
    try:
        if os.path.exists(temp_folder):
             shutil.rmtree(temp_folder)
    except Exception:
        pass

def check_package(package_name):
    """Verifica si un paquete requerido está instalado."""
    if package_name == "Pillow":
        return Image is not None
    elif package_name == "pyinstaller":
        # Check if 'pyinstaller' command is available
        try:
            # Usar sys.executable para mayor robustez
            subprocess.run([sys.executable, '-m', 'PyInstaller', '--version'], check=True, capture_output=True, text=True, timeout=5)
            return True
        except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
            try: # Intento alternativo por si está en el PATH sin -m
                 subprocess.run(['pyinstaller', '--version'], check=True, capture_output=True, text=True, timeout=5)
                 return True
            except:
                 return False
    return False

# ====================================================
# Funciones de Verificación de Dependencias (Consola)
# ====================================================

def check_all_requirements():
    """Verifica todos los requisitos y muestra el resultado en la consola/messagebox."""
    missing = []
        
    for pkg in REQUIRED_PACKAGES:
        if not check_package(pkg):
            missing.append(pkg)
            
    if missing:
        msg = f"ERROR CRÍTICO: Faltan paquetes de Python requeridos para la compilación: {', '.join(missing)}.\n\n"
        msg += "Por favor, instálalos manualmente usando el siguiente comando antes de continuar:\n"
        msg += f"pip install {' '.join(missing)}"
        messagebox.showerror("Error de Dependencias", msg)
        sys.exit(1)
        
    print(f"\n[OK] Verificación de dependencias completa. PyInstaller y Pillow listos.\n")
    return True


# ====================================================
# FUNCIONES DE COMPILACIÓN
# ====================================================

def compilar_script_macos(ruta_script, nombre_base, icono_path, output_folder, text_widget, progress_bar, on_finish=None):
    writer = GuiWriter(text_widget)
    # --- Asumo que PYTHON_SISTEMA y PYINSTALLER_PATH son válidos en Mac ---
    if not os.path.exists(PYTHON_SISTEMA) or not os.path.exists(PYINSTALLER_PATH):
        messagebox.showerror("Error", "Rutas de Python o PyInstaller no válidas para macOS.")
        if on_finish: on_finish(1); return

    if not os.path.exists(output_folder): os.makedirs(output_folder, exist_ok=True)
    
    carpeta_temp = os.path.join(output_folder, f"temp_{nombre_base}_{int(time.time())}_mac")
    os.makedirs(carpeta_temp, exist_ok=True)
    
    workpath = os.path.join(carpeta_temp, "build")
    specpath = carpeta_temp
    distpath = os.path.join(carpeta_temp, "dist") 
    
    comando = [
        PYTHON_SISTEMA,
        PYINSTALLER_PATH,
        "--noconfirm",
        "--windowed",
        "--onedir", # Usamos --onedir, que es más robusto para archivos adjuntos en Mac
        f"--workpath={workpath}",
        f"--specpath={specpath}",
        f"--distpath={distpath}",
        f"--name={nombre_base}",
        ruta_script
    ]
    
    if icono_path and os.path.exists(icono_path) and Path(icono_path).suffix.lower() == '.icns':
         comando.append(f"--icon={icono_path}")
         
    # 🟢 CORRECCIÓN DE ADJUNTOS EN MAC
    if AppGUI.instance and AppGUI.instance.adjuntos:
        sep = os.pathsep # ':' en Mac
        for f in AppGUI.instance.adjuntos:
             if os.path.exists(f):
                 abs_f = os.path.abspath(f)
                 dest_dir = Path(f).name # Usamos el nombre del archivo/carpeta como destino
                 # Formato SOURCE:DEST_FOLDER_NAME
                 comando.append(f"--add-data={abs_f}{sep}{dest_dir}")
    
    writer.write("=== [2%] Ejecutando PyInstaller para macOS (.app) ===\n")
    
    ret_code = 1
    try:
        progress_bar.start(50)
        proceso = subprocess.run(comando, check=False, capture_output=True, text=True)
        writer.write(proceso.stdout)
        writer.write(proceso.stderr)
        
        ret_code = proceso.returncode # Obtener el código de PyInstaller

        if ret_code == 0:
            writer.report_progress(90, "[90%] Moviendo archivo final y limpiando temporales...")
            ret_code = move_file_and_cleanup(carpeta_temp, None, ".app", output_folder, nombre_base, writer)
            if ret_code == 0:
                 play_snake_hiss() # Sonido de éxito
        else:
            writer.write(f"\n[ERROR] PyInstaller finalizó con código {proceso.returncode}.\n")
            ret_code = 1
    except Exception as e:
        writer.write(f"Error durante la compilación: {e}\n")
        ret_code = 1
    finally:
        progress_bar.stop()
        if on_finish: on_finish(ret_code)

def compilar_para_exe(ruta_script, nombre_base, icono_path, archivos_adjuntos, output_folder, text_widget, progress_bar, on_finish=None):
    writer = GuiWriter(text_widget)
    
    # 🟢 DUALIDAD: Asignación de comandos basada en el OS actual
    # ⚠️ CORRECCIÓN CLAVE: Usar sys.executable -m PyInstaller para máxima robustez en Win/Mac
    
    if not os.path.exists(output_folder): os.makedirs(output_folder, exist_ok=True)

    carpeta_temp = os.path.join(output_folder, f"temp_{nombre_base}_{int(time.time())}_exe")
    
    cmd = [
        sys.executable,  # Usa el intérprete de Python que ejecutó este script
        '-m', 'PyInstaller', # Lanza PyInstaller como un módulo, el método más robusto
        "--noconfirm",
        "--onefile", # Para .exe, usamos --onefile
        "--windowed",
        f"--workpath={carpeta_temp}/build", 
        f"--specpath={carpeta_temp}",        
        f"--distpath={carpeta_temp}/dist",   
        f"--name={nombre_base}",
        ruta_script
    ]

    if icono_path and os.path.exists(icono_path) and Path(icono_path).suffix.lower() == '.ico':
        cmd.append(f"--icon={icono_path}")

    # Archivos y carpetas adjuntas
    if archivos_adjuntos:
        sep = os.pathsep 
        for f in archivos_adjuntos:
            if os.path.exists(f):
                abs_f = os.path.abspath(f)
                dest_dir = Path(f).name 
                # En Windows, usamos el separador de directorio de Windows para abs_f
                if sys.platform.startswith("win"):
                    abs_f = abs_f.replace('\\', '/') 
                cmd.append(f"--add-data={abs_f}{os.pathsep}{dest_dir}")
                
    writer.write("=== [2%] Ejecutando PyInstaller para Windows (.exe) ===\n")

    def hilo():
        ret_code = 1
        try:
            progress_bar.start(50)
            # Usar Popen para capturar la salida en tiempo real
            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, bufsize=1)
        except Exception as e:
            writer.write(f"[ERROR al iniciar PyInstaller EXE]: {e}\n")
            progress_bar.stop()
            move_file_and_cleanup(carpeta_temp, None, ".exe", output_folder, nombre_base, writer) 
            if on_finish: on_finish(1); return

        def reader(pipe):
            for line in iter(pipe.readline, ''):
                if not line: break
                writer.write(line)
            pipe.close()

        t1 = threading.Thread(target=reader, args=(proc.stdout,), daemon=True)
        t2 = threading.Thread(target=reader, args=(proc.stderr,), daemon=True)
        t1.start(); t2.start()
        proc.wait()
        
        t1.join(); t2.join() 
        
        progress_bar.stop()
        ret_code = proc.returncode
        
        if ret_code == 0:
            writer.report_progress(90, "[90%] Moviendo archivo final y limpiando temporales...")
            ret_code = move_file_and_cleanup(carpeta_temp, None, ".exe", output_folder, nombre_base, writer)
        else:
            writer.write(f"\n[PyInstaller finalizó con código {ret_code}]\n")
            
        if ret_code == 0:
            play_snake_hiss() # <--- SONIDO DE ÉXITO PARA WINDOWS
            
        if on_finish: on_finish(ret_code)

    threading.Thread(target=hilo, daemon=True).start()


def run_compilar_in_thread(ruta_script, nombre_base, icono_path, output_folder, text_widget, progress_bar, on_finish=None):
    """Ejecuta la compilación en un hilo para mantener la UI responsiva."""
    def target():
        pw = GuiWriter(text_widget)
        orig_stdout = sys.stdout
        orig_stderr = sys.stderr
        ret = 1
        try:
            sys.stdout = pw
            sys.stderr = pw
            
            # DUALIDAD: Llama a la función de compilación correcta
            if sys.platform.startswith("darwin"):
                compilar_script_macos(ruta_py, nombre_salida, icono_mac, output, self.text_debug, self.progress, on_finish=on_finish)
            elif sys.platform.startswith("win"):
                compilar_para_exe(ruta_py, nombre_salida, icono_win, self.adjuntos, output, self.text_debug, self.progress, on_finish=on_finish)
            else:
                compilar_para_exe(ruta_py, nombre_salida, icono_win, self.adjuntos, output, self.text_debug, self.progress, on_finish=on_finish) 
            return 
        except Exception as e:
            pw.write(f"\n[EXCEPCIÓN en la compilación]: {e}\n")
            ret = 1
        finally:
            # Restaurar la salida
            sys.stdout = orig_stdout
            sys.stderr = orig_stderr
            if on_finish and ret != 0:
                try:
                    on_finish(ret)
                except Exception:
                    pass

    t = threading.Thread(target=target, daemon=True)
    t.start()


# ====================================================
# GUI principal
# ====================================================

class AppGUI:
    # Variable para almacenar la instancia 
    instance = None 

    def __init__(self, root):
        self.root = root
        self.root.title("") 
        
        AppGUI.instance = self # Almacenar la instancia actual
        
        style = ttk.Style()
        style.theme_use('default')

        default_font = ("Helvetica", DEFAULT_FONT_SIZE)
        header_font = ("Helvetica", HEADER_FONT_SIZE, "bold")
        big_font = ("Helvetica", BIG_FONT_SIZE)
        
        root.configure(bg=COLOR_BG)

        container = tk.Frame(root, bg=COLOR_BG, padx=int(12*SCALE_FACTOR), pady=int(12*SCALE_FACTOR))
        container.pack(fill=tk.BOTH, expand=True)
        
        # --- Header ---
        self.header = tk.Label(container, text=APP_TITLE, bg=COLOR_BG, fg=COLOR_FG, font=header_font)
        self.header.pack(fill=tk.X, anchor="n", pady=(0,int(10*SCALE_FACTOR)))

        # --- Formulario ---
        form = tk.Frame(container, bg=COLOR_BG)
        form.pack(fill=tk.X, pady=(0,int(10*SCALE_FACTOR)))
        
        form.grid_columnconfigure(0, weight=0) 
        form.grid_columnconfigure(1, weight=1) 
        form.grid_columnconfigure(2, weight=0) 

        FORM_BTN_BG = "#ffffff" 
        FORM_BTN_FG = "black"   

        # Row: archivo .py
        tk.Label(form, text="Archivo .py:", bg=COLOR_BG, fg=COLOR_FG, font=big_font).grid(row=0, column=0, sticky="w", pady=int(6*SCALE_FACTOR))
        self.entry_py = tk.Entry(form, font=default_font, width=ENTRY_WIDTH)
        self.entry_py.grid(row=0, column=1, padx=int(8*SCALE_FACTOR), sticky="ew")
        btn_browse_py = tk.Button(form, text="📂 Buscar", bg=FORM_BTN_BG, fg=FORM_BTN_FG, font=default_font, command=self.browse_py, width=int(10*SCALE_FACTOR))
        btn_browse_py.grid(row=0, column=2, padx=int(6*SCALE_FACTOR), sticky="ew")

        # Row: nombre de salida
        tk.Label(form, text="Nombre de salida:", bg=COLOR_BG, fg=COLOR_FG, font=big_font).grid(row=1, column=0, sticky="w", pady=int(6*SCALE_FACTOR))
        self.entry_name = tk.Entry(form, font=default_font, width=ENTRY_WIDTH)
        self.entry_name.grid(row=1, column=1, padx=int(8*SCALE_FACTOR), sticky="ew")
        btn_default_name = tk.Button(form, text="Nombre Default", bg=FORM_BTN_BG, fg=FORM_BTN_FG, font=default_font, command=self.set_default_name, width=int(10*SCALE_FACTOR))
        btn_default_name.grid(row=1, column=2, padx=int(6*SCALE_FACTOR), sticky="ew")

        # Row: icono/imagen
        tk.Label(form, text="Icono (.png/.jpg/.ico/.icns):", bg=COLOR_BG, fg=COLOR_FG, font=big_font).grid(row=2, column=0, sticky="w", pady=int(6*SCALE_FACTOR))
        self.entry_icon = tk.Entry(form, font=default_font, width=ENTRY_WIDTH)
        self.entry_icon.grid(row=2, column=1, padx=int(8*SCALE_FACTOR), sticky="ew")
        tk.Button(form, text="🖼️ Buscar", bg=FORM_BTN_BG, fg=FORM_BTN_FG, font=default_font, command=self.browse_icon, width=int(10*SCALE_FACTOR)).grid(row=2, column=2, padx=int(6*SCALE_FACTOR), sticky="ew")

        # Row: ruta final
        tk.Label(form, text="Carpeta destino:", bg=COLOR_BG, fg=COLOR_FG, font=big_font).grid(row=3, column=0, sticky="w", pady=int(6*SCALE_FACTOR))
        self.entry_output = tk.Entry(form, font=default_font, width=ENTRY_WIDTH)
        self.entry_output.grid(row=3, column=1, padx=int(8*SCALE_FACTOR), sticky="ew")
        tk.Button(form, text="📁 Carpeta", bg=FORM_BTN_BG, fg=FORM_BTN_FG, font=default_font, command=self.browse_output, width=int(10*SCALE_FACTOR)).grid(row=3, column=2, padx=int(6*SCALE_FACTOR), sticky="ew")
        self.entry_output.insert(0, CARPETA_SALIDA_DEFECTO)
        
        # --- Adjuntos Box ---
        adjuntos_frame = tk.Frame(container, bg=COLOR_BOX_BG, padx=int(10*SCALE_FACTOR), pady=int(10*SCALE_FACTOR))
        adjuntos_frame.pack(fill=tk.X, pady=(int(10*SCALE_FACTOR), int(10*SCALE_FACTOR)))
        
        tk.Label(adjuntos_frame, text="Archivos/Carpetas Necesarias (Adjuntos):", bg=COLOR_BOX_BG, fg=COLOR_FG, font=big_font).pack(anchor="w", pady=(0, 5))
        
        adjuntos_list_frame = tk.Frame(adjuntos_frame, bg=COLOR_BOX_BG)
        adjuntos_list_frame.pack(fill=tk.X, expand=True)
        
        self.listbox_adjuntos = tk.Listbox(adjuntos_list_frame, height=int(5*SCALE_FACTOR), bg=COLOR_BG, fg=COLOR_FG, font=default_font, selectmode=tk.EXTENDED)
        self.listbox_adjuntos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5))
        adjuntos_scrollbar = ttk.Scrollbar(adjuntos_list_frame, command=self.listbox_adjuntos.yview)
        adjuntos_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox_adjuntos.config(yscrollcommand=adjuntos_scrollbar.set)
        
        adjuntos_btns_frame = tk.Frame(adjuntos_frame, bg=COLOR_BOX_BG, pady=5)
        adjuntos_btns_frame.pack(fill=tk.X)
        
        inner_btn_frame = tk.Frame(adjuntos_btns_frame, bg=COLOR_BOX_BG)
        inner_btn_frame.pack(expand=True)
        
        btn_width = int(140 * SCALE_FACTOR)
        btn_height = int(35 * SCALE_FACTOR)
        
        # Botones de adjuntos con la funcionalidad de añadir archivos/carpetas
        self.btn_add_adjuntos = RoundedButton(inner_btn_frame, text="➕ Añadir", command=self.browse_adjuntos, 
                      width=btn_width, height=btn_height, 
                      fill=COLOR_ACCENT_LIGHT, fg="black")
        self.btn_add_adjuntos.pack(side=tk.LEFT, padx=int(8*SCALE_FACTOR))
        
        self.btn_remove_adjunto = RoundedButton(inner_btn_frame, text="🗑️ Eliminar", command=self.remove_selected_adjunto, 
                      width=btn_width, height=btn_height, 
                      fill=COLOR_ACCENT_LIGHT, fg="black")
        self.btn_remove_adjunto.pack(side=tk.LEFT, padx=int(8*SCALE_FACTOR))
        
        self.btn_clear_adjuntos = RoundedButton(inner_btn_frame, text="🧹 Borrar Todo", command=self.clear_adjuntos, 
                      width=btn_width, height=btn_height, 
                      fill=COLOR_ACCENT_LIGHT, fg="black")
        self.btn_clear_adjuntos.pack(side=tk.LEFT, padx=int(8*SCALE_FACTOR))

        self.adjuntos = [] # Lista de rutas completas a archivos y/o carpetas

        # --- Buttons: Compilar (todo el ancho) ---
        self.btn_compile = tk.Button(container, 
                                     text=f"COMPILAR PARA {'MAC (.APP)' if sys.platform.startswith('darwin') else 'WINDOWS (.EXE)'}", 
                                     bg=COLOR_BTN_COMPILE, fg="black", 
                                     font=("Helvetica", BIG_FONT_SIZE + 2, "bold"), command=self.start_compile_all, 
                                     height=int(2*SCALE_FACTOR))
        self.btn_compile.pack(fill=tk.X, pady=(int(10*SCALE_FACTOR), int(10*SCALE_FACTOR)))

        # --- Barra de progreso y zona debug ---
        progress_frame = tk.Frame(container, bg=COLOR_BG, pady=int(8*SCALE_FACTOR), padx=0)
        progress_frame.pack(fill=tk.X, expand=False)

        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="indeterminate") 
        self.progress.pack(fill=tk.X, pady=(int(4*SCALE_FACTOR), int(8*SCALE_FACTOR)))

        # Debug text area
        debug_frame = tk.Frame(container, bg=COLOR_BOX_BG, pady=5, padx=5)
        debug_frame.pack(fill=tk.BOTH, expand=True)
        # 🟢 CORRECTO: Fondo Oscuro para la salida de la compilación
        self.text_debug = tk.Text(debug_frame, height=int(10*SCALE_FACTOR), 
                                   bg="#021827", fg=COLOR_FG, # Fondo Oscuro, Texto Blanco
                                   font=("Courier", DEBUG_FONT_SIZE), wrap="none")
        self.text_debug.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(debug_frame, command=self.text_debug.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_debug.config(yscrollcommand=scrollbar.set)
        
        # Botón Limpiar Debug
        tk.Button(container, text="Limpiar Debug", bg=COLOR_ACCENT_DARK, fg=COLOR_BOX_BG, font=default_font, command=self.clear_debug).pack(anchor="w", pady=(5,0))

        # Limpiar cualquier ícono temporal de una ejecución previa
        cleanup_temp_icons()
        
        self.root.update_idletasks() 
        self.adjust_and_center_window(self.root)

    # ------------------ METODOS DE UI Y UTILIDAD ------------------

    def adjust_and_center_window(self, root):
        """Calcula el ancho actual de la ventana y la centra."""
        window_width = root.winfo_reqwidth()
        window_height = root.winfo_reqheight()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        root.resizable(False, False)

    def browse_py(self):
        ruta = filedialog.askopenfilename(title="Seleccionar script Python", filetypes=[("Archivos Python", "*.py")])
        if ruta:
            self.entry_py.delete(0, tk.END)
            self.entry_py.insert(0, ruta)

    def browse_icon(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar Icono (.png, .jpg, .ico, .icns)", 
            filetypes=[
                ("Archivos de Imagen", "*.icns *.ico *.png *.jpg *.jpeg"), 
                ("Todos", "*.*")
            ]
        )
        if ruta:
            self.entry_icon.delete(0, tk.END)
            self.entry_icon.insert(0, ruta)
            
    def set_default_name(self):
        """Genera un nombre de salida por defecto basado en el script .py."""
        ruta_py = self.entry_py.get().strip()
        if not ruta_py:
            messagebox.showwarning("Atención", "Selecciona primero un archivo .py válido para generar el nombre por defecto.")
            return

        base_name = os.path.splitext(os.path.basename(ruta_py))[0]
        output_folder = self.entry_output.get().strip() or CARPETA_SALIDA_DEFECTO
        
        i = 1
        default_name = base_name
        # Comprobar la existencia en la carpeta de salida
        while (os.path.exists(os.path.join(output_folder, f"{default_name}.app")) or 
               os.path.exists(os.path.join(output_folder, f"{default_name}.exe"))):
            default_name = f"{base_name}_{i}"
            i += 1
        
        self.entry_name.delete(0, tk.END)
        self.entry_name.insert(0, default_name)

    def browse_adjuntos(self):
        """Permite multiselección de archivos y/o la selección de una carpeta."""
        
        rutas_archivos = filedialog.askopenfilenames(title="Seleccionar Archivos Adjuntos (Ctrl+Click para multiselección)")
        ruta_carpeta = filedialog.askdirectory(title="Seleccionar Carpeta Adjunta (Opcional)")
        
        rutas = list(rutas_archivos)
        if ruta_carpeta:
            rutas.append(ruta_carpeta)

        if rutas:
            for r in rutas:
                if r not in self.adjuntos:
                    self.adjuntos.append(r)
                    
                    display_name = os.path.basename(r)
                    if os.path.isdir(r):
                         display_name = f"[CARPETA] {display_name}"
                    
                    self.listbox_adjuntos.insert(tk.END, display_name)

    def remove_selected_adjunto(self):
        selected_indices = self.listbox_adjuntos.curselection()
        if not selected_indices: return
        
        for i in reversed(selected_indices):
            del self.adjuntos[i]
            self.listbox_adjuntos.delete(i)

    def clear_adjuntos(self):
        self.adjuntos = []
        self.listbox_adjuntos.delete(0, tk.END)

    def browse_output(self):
        carpeta = filedialog.askdirectory(title="Selecciona carpeta destino", initialdir=self.entry_output.get().strip() or CARPETA_SALIDA_DEFECTO)
        if carpeta:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, carpeta)

    def clear_debug(self):
        self.text_debug.delete(1.0, tk.END)


    def start_compile_all(self):
        ruta_py = self.entry_py.get().strip()
        if not ruta_py or not os.path.exists(ruta_py):
            messagebox.showwarning("Atención", "Selecciona primero un archivo .py válido para compilar.")
            return

        nombre_salida = self.entry_name.get().strip()
        if not nombre_salida:
            nombre_salida = os.path.splitext(os.path.basename(ruta_py))[0]

        icono_original = self.entry_icon.get().strip()
        output = self.entry_output.get().strip() or CARPETA_SALIDA_DEFECTO
        
        # Desactivar UI
        self.btn_compile.configure(state="disabled")
        for child in self.root.winfo_children():
            for inner_child in child.winfo_children():
                 if isinstance(inner_child, RoundedButton):
                    inner_child.set_state("disabled")
                 elif isinstance(inner_child, tk.Frame):
                    for sub_child in inner_child.winfo_children():
                         if isinstance(sub_child, tk.Frame):
                             for btn in sub_child.winfo_children():
                                 if isinstance(btn, tk.Button):
                                     btn.configure(state="disabled")

        # Procesar Ícono antes de compilar
        icono_mac = None
        icono_win = None
        
        if icono_original and os.path.exists(icono_original):
            if process_icon_file(icono_original, output): 
                
                path_obj = Path(icono_original)
                base_name = path_obj.stem
                temp_folder = os.path.join(output, "temp_icons")
                
                # Intentar usar el archivo .icns generado o el original si era .icns
                temp_icns = os.path.join(temp_folder, f"{base_name}.icns")
                if path_obj.suffix.lower() == '.icns':
                     icono_mac = icono_original
                elif os.path.exists(temp_icns): 
                     icono_mac = temp_icns
                
                # Intentar usar el archivo .ico generado o el original si era .ico
                temp_ico = os.path.join(temp_folder, f"{base_name}.ico")
                if path_obj.suffix.lower() == '.ico':
                     icono_win = icono_original
                elif os.path.exists(temp_ico): 
                     icono_win = temp_ico


        def runner():
            # Función para reactivar la UI
            def reactivate_ui():
                try:
                    self.btn_compile.configure(state="normal")
                    for child in self.root.winfo_children():
                        for inner_child in child.winfo_children():
                            if isinstance(inner_child, RoundedButton):
                                inner_child.set_state("normal")
                            elif isinstance(inner_child, tk.Frame):
                                for sub_child in inner_child.winfo_children():
                                    if isinstance(sub_child, tk.Frame):
                                        for btn in sub_child.winfo_children():
                                            if isinstance(btn, tk.Button):
                                                btn.configure(state="normal")
                except Exception:
                    pass

            pw = GuiWriter(self.text_debug)
            orig_stdout = sys.stdout
            orig_stderr = sys.stderr
            ret = 1
            try:
                sys.stdout = pw
                sys.stderr = pw

                # 🟢 DUALIDAD: Llama a la función de compilación correcta
                if sys.platform.startswith("darwin"):
                    compilar_script_macos(ruta_py, nombre_salida, icono_mac, output, self.text_debug, self.progress, on_finish=on_finish)
                elif sys.platform.startswith("win"):
                    compilar_para_exe(ruta_py, nombre_salida, icono_win, self.adjuntos, output, self.text_debug, self.progress, on_finish=on_finish)
                else:
                    compilar_para_exe(ruta_py, nombre_salida, icono_win, self.adjuntos, output, self.text_debug, self.progress, on_finish=on_finish) 
                
            except Exception as e:
                pw.write(f"\n[EXCEPCIÓN en la compilación]: {e}\n")
                ret = 1
            finally:
                # Limpiar Íconos Temporales y Reactivar UI
                sys.stdout = orig_stdout
                sys.stderr = orig_stderr
                cleanup_temp_icons() 
                self.root.after(0, reactivate_ui)


        # 🟢 CORRECCIÓN DEL ERROR CRÍTICO:
        # Definir la función on_finish aquí para que runner la pueda usar,
        # y luego iniciar el hilo.
        def on_finish(retcode):
            pass # Solo una función placeholder, ya que runner() manejará el resultado en el finally.

        threading.Thread(target=runner, daemon=True).start()

# ====================================================
# Punto de entrada
# ====================================================
def main():
    root = tk.Tk()
    
    # 🟢 1. Verificar requisitos (imprimir a consola/salir si falla)
    check_all_requirements() 
    
    # 🟢 2. Iniciar la GUI principal (ya está garantizado que las dependencias están)
    AppGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()