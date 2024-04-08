import wmi
import tkinter as tk
from PIL import Image, ImageTk

def get_installed_applications():
    installed_apps = []
    try:
        # Conectarse al servicio WMI de Windows
        wmi_service = wmi.WMI()
        # Consultar los programas instalados
        programs = wmi_service.Win32_Product()
        installed_apps = [(program.Caption, program.Version, program.InstallLocation) for program in programs]
    except Exception as e:
        print(f"Error: {e}")
    return installed_apps

def display_installed_apps():
    global y_position, card_height
    
    apps = get_installed_applications()
    if not apps:
        print("No se encontraron aplicaciones instaladas.")
    else:
        for app_name, app_version, app_path in apps:
            card_frame = tk.Frame(canvas, relief=tk.GROOVE, borderwidth=2)
            canvas.create_window((0, y_position), window=card_frame, anchor="nw")
            y_position += card_height
            
            # Label para el nombre y versión de la aplicación
            app_label = tk.Label(card_frame, text=f"{app_name} - {app_version}")
            app_label.pack(padx=5, pady=5, anchor="w")
            
            # Cargar la imagen de la aplicación si está disponible
            try:
                image_path = f"{app_path}\\{app_name}.png"
                img = Image.open(image_path)
                img.thumbnail((100, 100))
                img = ImageTk.PhotoImage(img)
                img_label = tk.Label(card_frame, image=img)
                img_label.image = img
                img_label.pack(padx=5, pady=5, anchor="w")
            except Exception as e:
                print(f"No se pudo cargar la imagen para {app_name}: {e}")

# Configuración de la ventana principal
root = tk.Tk()
root.title("Aplicaciones Instaladas")

# Configuración del área de desplazamiento
canvas = tk.Canvas(root)
scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollbar.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)
canvas.configure(yscrollcommand=scrollbar.set)

# Capturar el evento de desplazamiento con la rueda del ratón
canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

# Espacio en blanco para separar el botón de las tarjetas
tk.Label(root).pack()

# Frame interno para contener las tarjetas
frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=frame, anchor="nw")

# Lógica para ajustar el tamaño del canvas cuando se agregan tarjetas
def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

frame.bind("<Configure>", on_frame_configure)

# Botón para mostrar las aplicaciones
show_apps_button = tk.Button(root, text="Buscar Aplicaciones", command=display_installed_apps)
show_apps_button.pack(pady=5)

# Inicializar variables para el posicionamiento de las tarjetas
y_position = 0
card_height = 120

root.mainloop()