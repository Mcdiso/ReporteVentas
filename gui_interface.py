import tkinter as tk
from tkinter import Image, ttk, messagebox
from tkcalendar import DateEntry
from db_connection import get_databases, get_agents
from report_generation import generate_partida_report, generate_global_report, generate_client_report
from styles import apply_background, apply_styles, style_label, style_button, style_combobox
import os
from tkinter import PhotoImage
from PIL import Image, ImageTk 

def create_report_type_selection_window(parent_window, database, agent, start_date, end_date):
    """Crear la interfaz para seleccionar el tipo de reporte"""
    parent_window.destroy()  # Cerramos la ventana anterior

    window = tk.Tk()
    window.title("Seleccionar Tipo de Reporte")
    window.geometry("600x400")
    window.resizable(False, False)

    # Fondo degradado personalizado
    canvas = tk.Canvas(window, width=600, height=400, highlightthickness=0)
    canvas.pack(fill="both", expand=True)

    # Degradado
    for i in range(400):
        color = f"#{int(240 - (i / 400) * 100):02X}{int(240 - (i / 400) * 20):02X}{255:02X}"
        canvas.create_line(0, i, 600, i, fill=color)

    # Marco principal
    frame = tk.Frame(canvas, bg="white", padx=20, pady=20, relief="groove", bd=3)
    frame.place(relx=0.5, rely=0.5, anchor="center")

    # Título de la ventana
    title_label = tk.Label(frame, text="Seleccione el Tipo de Reporte", font=("Arial", 16, 'bold'), bg="white", fg="#2F3C48")
    title_label.grid(row=0, column=0, columnspan=2, pady=10)

    # Información de contexto
    info_label = tk.Label(frame, text=f"Base de Datos: {database}\nAgente: {agent}\nRango: {start_date} a {end_date}", 
                          font=("Arial", 12), bg="white", fg="#4A90E2", justify="left")
    info_label.grid(row=1, column=0, columnspan=2, pady=10)

    # Opciones de reporte
    def generate_partida():
        generate_partida_report(database, agent, start_date, end_date)
        messagebox.showinfo("Reporte Generado", "El reporte de partidas se generó exitosamente.")

    def generate_global():
        generate_global_report(database, agent, start_date, end_date)
        messagebox.showinfo("Reporte Generado", "El reporte global se generó exitosamente.")

    def generate_client():
        generate_client_report(database, agent, start_date, end_date)
        messagebox.showinfo("Reporte Generado", "El reporte de clientes se generó exitosamente.")

    # Botones de selección
    partida_button = tk.Button(frame, text="Reporte de Partidas", command=generate_partida, font=("Arial", 12), bg="#4A90E2", fg="white", width=20, relief="flat")
    partida_button.grid(row=2, column=0, pady=10, padx=10)

    global_button = tk.Button(frame, text="Reporte Global", command=generate_global, font=("Arial", 12), bg="#50B83C", fg="white", width=20, relief="flat")
    global_button.grid(row=2, column=1, pady=10, padx=10)

    client_button = tk.Button(frame, text="Reporte de Clientes", command=generate_client, font=("Arial", 12), bg="#FFB545", fg="white", width=20, relief="flat")
    client_button.grid(row=3, column=0, columnspan=2, pady=10)

    # Firma
    signature = tk.Label(window, text="By: Ing. Jesús Mena", font=("Arial", 10), fg="gray", bg="white")
    signature.place(relx=0.97, rely=1, anchor="se")

    window.mainloop()
    
def create_interface():
    """Crear la interfaz de usuario"""
    window = tk.Tk()
    window.title("Selección de Base de Datos")
    window.geometry("600x400")
    window.resizable(False, False)  # Evitar que se redimensione la ventana
    
    # Ajustar opacidad (en un rango de 0.0 a 1.0)
    window.attributes("-alpha", 0.9)  # 0.8 es un valor semi-transparente
    
    # Dividir la ventana en tres partes
    top_frame = tk.Frame(window, bg="#FFB6C1", height=100)  # Rojo claro en la parte superior
    top_frame.pack(fill="x", side="top")

    middle_frame = tk.Frame(window, bg="white", height=200)  # Blanco para la parte central (con los controles)
    middle_frame.pack(fill="both", expand=True)

    bottom_frame = tk.Frame(window, bg="#ADD8E6", height=100)  # Azul claro en la parte inferior
    bottom_frame.pack(fill="x", side="bottom")

    # Aplicar el logo en la parte superior central
    logo_path = "logo_empresa.jpg"  # Asegúrate de que la ruta del logo sea correcta
    try:
        logo = Image.open(logo_path)  # Usamos Pillow para abrir la imagen
        logo = logo.resize((200, 50))  # Ajustamos el tamaño del logo
        logo = ImageTk.PhotoImage(logo)  # Convertimos la imagen a un formato compatible con Tkinter
    except Exception as e:
        print(f"Error al cargar la imagen: {e}")
        logo = None  # Si hay un error, no mostramos la imagen

    if logo:
        logo_label = tk.Label(top_frame, image=logo, bg="#FFB6C1")  # Colocamos el logo en el top_frame
        logo_label.image = logo  # Necesario para mantener una referencia al logo
        logo_label.pack(pady=10)  # Aseguramos que el logo esté centrado y con algo de espacio

    # Crear un marco para centrar los elementos dentro de la sección central
    frame = tk.Frame(middle_frame, bg="white")
    frame.place(relx=0.5, rely=0.5, anchor="center")  # Centra el marco en la ventana

    # Etiqueta de texto centrada
    label = tk.Label(frame, text="Selecciona una Empresa:", font=("Arial", 14), bg="lightblue", fg="black")
    label.grid(row=0, column=0, pady=10, padx=10)

    # Crear el ComboBox centrado
    databases = get_databases()
    if not databases:
        messagebox.showerror("Error", "No se encontraron Empresas :c.")
        return

    db_combobox = ttk.Combobox(frame, values=databases, width=35, font=("Arial", 12))
    db_combobox.grid(row=1, column=0, pady=10)

    def on_db_selected():
        """Acción cuando se selecciona la base de datos"""
        selected_db = db_combobox.get()
        if selected_db:
            messagebox.showinfo("Base de datos seleccionada", f"Has seleccionado: {selected_db}")
            window.destroy()
            create_agent_and_date_interface(selected_db)
        else:
            messagebox.showwarning("Advertencia", "Por favor selecciona una base de datos.")

    # Botón de seleccionar centrado
    select_button = tk.Button(frame, text="Seleccionar", command=on_db_selected, font=("Arial", 12), width=20)
    select_button.grid(row=2, column=0, pady=10)

    # Firma en la parte inferior derecha
    signature = tk.Label(window, text="By: Ing. Jesús Mena", font=("Arial", 10), fg="gray", bg="white")
    signature.place(relx=0.97, rely=1, anchor="se")  # Firma un poco más abajo

    window.mainloop()

def create_agent_and_date_interface(database):
    """Crear la interfaz para seleccionar el agente de ventas y las fechas"""
    window = tk.Tk()
    window.title(f"Seleccionar Agente de Venta y Fechas - Base de Datos: {database}")
    window.geometry("720x450")  # Incrementamos el tamaño para mayor visibilidad
    window.resizable(False, False)

    # Fondo de la ventana
    window.configure(bg="#2F3C48")

    # Marco principal
    frame = tk.Frame(window, bg="#D1D9E6", padx=20, pady=20)
    frame.pack(pady=20, padx=20, fill="both", expand=True)

    # Etiqueta de título
    title_label = tk.Label(frame, text="Selección de Agente de Venta y Fechas", font=("Arial", 16, 'bold'), bg="#D1D9E6", fg="#2F3C48")
    title_label.grid(row=0, column=0, columnspan=2, pady=10)

    # Agente de ventas
    agent_label = tk.Label(frame, text="Selecciona un Agente de Venta:", font=("Arial", 12), bg="#D1D9E6", fg="#2F3C48")
    agent_label.grid(row=1, column=0, sticky="w", padx=10, pady=5)
    
    agents = get_agents(database)
    if not agents:
        messagebox.showerror("Error", "No se encontraron agentes de ventas.")
        return

    agent_combobox = ttk.Combobox(frame, values=agents, width=40, font=("Arial", 12))  # Ancho ajustado
    agent_combobox.grid(row=1, column=1, pady=5, padx=10)

    # Fechas
    date_label = tk.Label(frame, text="Selecciona las Fechas:", font=("Arial", 12), bg="#D1D9E6", fg="#2F3C48")
    date_label.grid(row=2, column=0, columnspan=2, pady=10)

    start_date_label = tk.Label(frame, text="Fecha de inicio:", font=("Arial", 12), bg="#D1D9E6", fg="#2F3C48")
    start_date_label.grid(row=3, column=0, sticky="w", padx=10)

    start_date_entry = DateEntry(frame, width=18, background='darkblue', foreground='white', font=("Arial", 12), date_pattern='dd-mm-yyyy')
    start_date_entry.grid(row=3, column=1, pady=5, padx=10)

    end_date_label = tk.Label(frame, text="Fecha de fin:", font=("Arial", 12), bg="#D1D9E6", fg="#2F3C48")
    end_date_label.grid(row=4, column=0, sticky="w", padx=10)

    end_date_entry = DateEntry(frame, width=18, background='darkblue', foreground='white', font=("Arial", 12), date_pattern='dd-mm-yyyy')
    end_date_entry.grid(row=4, column=1, pady=5, padx=10)

    def on_generate_report():
        """Generar el reporte usando los datos ingresados"""
        selected_agent = agent_combobox.get()
        start_date = start_date_entry.get_date()
        end_date = end_date_entry.get_date()

        if not selected_agent or not start_date or not end_date:
            messagebox.showwarning("Advertencia", "Por favor complete todos los campos.")
            return

        start_date_str = start_date.strftime('%d-%m-%Y')
        end_date_str = end_date.strftime('%d-%m-%Y')

        # Llamar a la función para crear el tipo de reporte
        create_report_type_selection_window(window, database, selected_agent, start_date_str, end_date_str)

    # Botón para generar el reporte
    generate_button = tk.Button(frame, text="Generar Reporte", command=on_generate_report, font=("Arial", 12), width=20, bg="#4A90E2", fg="white", relief="flat", bd=2)
    generate_button.grid(row=5, column=0, columnspan=2, pady=20)

    # Firma en la parte inferior del marco
    signature = tk.Label(frame, text="By: Ing. Jesús Mena", font=("Arial", 10), fg="gray", bg="#D1D9E6")
    signature.grid(row=6, column=0, columnspan=2, pady=10)

    window.mainloop()

create_interface()