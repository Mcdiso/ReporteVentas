import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk, ImageDraw

# Colores predeterminados
BACKGROUND_COLOR = "#f5f5f5"
BUTTON_COLOR = "#4CAF50"
BUTTON_HOVER_COLOR = "#45a049"
LABEL_COLOR = "#333333"
FONT = ("Arial", 12)

def apply_styles(window):
    """Aplicar estilos generales a la ventana"""
    window.config(bg=BACKGROUND_COLOR)

def style_label(label):
    """Estilizar etiquetas"""
    label.config(bg=BACKGROUND_COLOR, fg=LABEL_COLOR, font=("Arial", 14, "bold"))

def style_button(button):
    """Estilizar botones"""
    button.config(bg=BUTTON_COLOR, fg="white", relief="raised", font=FONT)
    button.bind("<Enter>", lambda e: button.config(bg=BUTTON_HOVER_COLOR))  # Cambiar color al pasar el ratón
    button.bind("<Leave>", lambda e: button.config(bg=BUTTON_COLOR))  # Restablecer el color

def style_combobox(combobox):
    """Estilizar el combobox"""
    combobox.config(width=40, font=FONT)
    combobox["state"] = "readonly"  # Solo lectura, no editable

def apply_background(window, logo_path, logo_size):
    """Aplica un fondo de color y coloca el logo en la parte superior."""
    from PIL import Image, ImageTk

    # Fondo de color
    window.configure(bg="white")

    # Logo en la parte superior
    try:
        logo = Image.open(logo_path)
        logo = logo.resize(logo_size, Image.Resampling.LANCZOS)
        logo = ImageTk.PhotoImage(logo)
        logo_label = tk.Label(window, image=logo, bg="white")
        logo_label.image = logo
        logo_label.place(x=0, y=0, relwidth=1, height=logo_size[1])
    except Exception as e:
        print(f"Error cargando el logo: {e}")

def style_combobox(combobox):
    """
    Aplica estilos al combobox.
    """
    style = ttk.Style()
    style.theme_use('clam')  # Cambia el tema si es necesario
    style.configure("TCombobox", fieldbackground="white", background="lightblue", font=("Arial", 12))

def style_button(button):
    """
    Aplica estilos al botón.
    """
    button.config(
        font=("Arial", 12),
        bg="lightblue",
        fg="black",
        activebackground="darkblue",
        activeforeground="white",
        relief="raised"
    )

def apply_background(window, logo_path, size):
    """
    Aplica un fondo difuminado con un logo al final de la ventana.
    """
    from tkinter import Canvas  # Importar aquí para evitar conflictos circulares
    canvas = Canvas(window, width=size[0], height=size[1])
    canvas.pack(fill="both", expand=True)

    # Crear fondo con degradado
    bg_image = Image.new("RGB", size, "black")
    draw = ImageDraw.Draw(bg_image)
    for y in range(size[1]):
        color = (
            int(255 * (y / size[1])),  # Gradiente de rojo
            0,                         # Sin verde
            int(255 * ((size[1] - y) / size[1]))  # Gradiente de azul
        )
        draw.line([(0, y), (size[0], y)], fill=color)

    bg_photo = ImageTk.PhotoImage(bg_image)
    canvas.create_image(0, 0, image=bg_photo, anchor="nw")

    # Colocar logo
    try:
        logo = Image.open(logo_path)
        logo_width = 300  # Ajusta el ancho del logo
        logo_height = 100  # Ajusta el alto del logo (puedes personalizarlo)
        logo = logo.resize((logo_width, logo_height), Image.Resampling.LANCZOS)
        logo_photo = ImageTk.PhotoImage(logo)
        canvas.create_image(size[0] // 2, size[1] - 100, image=logo_photo, anchor="center")
        canvas.logo_image = logo_photo  # Mantener referencia para evitar que se recoja por el GC
    except FileNotFoundError:
        print(f"Logo no encontrado en {logo_path}")

    return canvas

def style_combobox(combobox):
    """
    Aplica estilos al combobox.
    """
    style = ttk.Style()
    style.theme_use('clam')  # Cambia el tema si es necesario
    style.configure("TCombobox", fieldbackground="white", background="lightblue", font=("Arial", 12))

def style_button(button):
    """
    Aplica estilos al botón.
    """
    button.config(
        font=("Arial", 12),
        bg="lightblue",
        fg="black",
        activebackground="darkblue",
        activeforeground="white",
        relief="raised"
    )

def style_label(label):
    """
    Aplica estilos a las etiquetas (labels).
    """
    label.config(
        font=("Arial", 14),
        bg="#add8e6",  # Azul claro
        fg="black"
    )
