import tkinter as tk
from tkinter import filedialog, Label, Button
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from openpyxl.styles import Font
from PIL import Image, ImageTk



def compare_excels():
    if not ref_file_path or not new_file_path:
        status_label.config(text="Veuillez sélectionner les fichiers nécessaires.", fg="red")
        return

    status_text = f"Référence: {ref_file_path}\nNouveau: {new_file_path}"
    if output_dir:
        status_text += f"\nDossier de sortie: {output_dir}"

    status_label.config(text=status_text, fg="black")

from tkinter import messagebox

def select_ref_file():
    global ref_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    if not file_path.endswith((".xlsx", ".xls")):
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel valide (*.xlsx, *.xls)")
        return

    ref_file_path = file_path
    ref_label.config(text=f"Référence: {ref_file_path}")

def select_new_file():
    global new_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    if not file_path.endswith((".xlsx", ".xls")):
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel valide (*.xlsx, *.xls)")
        return

    new_file_path = file_path
    new_label.config(text=f"Nouveau: {new_file_path}")


def select_output_dir():
    global output_dir
    output_dir = filedialog.askdirectory()
    output_label.config(text=f"Dossier de sortie: {output_dir}")


root = tk.Tk()
root.title("Excel Comparator - Yazaki")
root.geometry("350x350")

# window_width = 300
# window_height = 300

# # Obtenir les dimensions de l'écran
# screen_width = root.winfo_screenwidth()
# screen_height = root.winfo_screenheight()

# # Calculer la position x et y
# x_position = (screen_width - window_width) // 2
# y_position = (screen_height - window_height) // 2

# # Appliquer la position centrée
# root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

logo_path = "yazaki_logo.png"  
try:
    img = Image.open(logo_path)
    img = img.resize((150, 50), Image.LANCZOS)
    logo = ImageTk.PhotoImage(img)
    root.logo = logo
    logo_label = Label(root, image=logo)
    logo_label.pack(pady=10)
except:
    logo_label = Label(root, text="[Logo Yazaki]", font=("Arial", 14, "bold"))
    logo_label.pack(pady=10)


ref_file_path = ""
new_file_path = ""
output_dir = ""

ref_button = Button(root, text="Sélectionner le fichier de référence", command=select_ref_file)
ref_button.pack(pady=5)
ref_label = Label(root, text="Référence: Non sélectionné", wraplength=400)
ref_label.pack()

new_button = Button(root, text="Sélectionner le fichier nouveau", command=select_new_file)
new_button.pack(pady=5)
new_label = Label(root, text="Nouveau: Non sélectionné", wraplength=400)
new_label.pack()

output_button = Button(root, text="Sélectionner le dossier de sortie", command=select_output_dir)
output_button.pack(pady=5)
output_label = Label(root, text="Dossier de sortie: Non sélectionné", wraplength=400)
output_label.pack()

compare_button = Button(root, text="Comparer", command=compare_excels)
compare_button.pack(pady=20)

status_label = Label(root, text="", fg="black")
status_label.pack()

tk.mainloop()