import tkinter as tk
from tkinter import filedialog, Label, Button
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from openpyxl.styles import Font
from PIL import Image, ImageTk

def add_comment(old_value, new_value):
    return f"Old: {old_value}\nNew: {new_value}"

def compare_excels():
    if not ref_file_path or not new_file_path or not output_dir:
        status_label.config(text="Veuillez sélectionner les fichiers et le dossier de sortie.", fg="red")
        return
    
    output_path = f"{output_dir}/output.xlsx"
    try:
        df_old_version = pd.read_excel(ref_file_path)
        df_new_version = pd.read_excel(new_file_path)

        if df_old_version.shape != df_new_version.shape:
            status_label.config(text="Erreur: Les fichiers ont des tailles différentes.", fg="red")
            return 
        
        wb_new = load_workbook(new_file_path)
        ws_new = wb_new.active

        if "Summary" in wb_new.sheetnames:
            wb_new.remove(wb_new["Summary"])
        ws_summary = wb_new.create_sheet("Summary")

        ws_summary.append(["Row", "Column", "Old Value", "New Value"])

        red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        font = Font(color='FFFFFF')

        for row_index in range(0, df_new_version.shape[0]):
            for col_index in range(1, df_new_version.shape[1]):
                old_value = df_old_version.iloc[row_index, col_index]
                ref_value = df_new_version.iloc[row_index, col_index]
                
                if pd.isna(old_value) and pd.isna(ref_value):
                    continue
                elif old_value != ref_value or (pd.isna(old_value) or pd.isna(ref_value)):
                    cell = ws_new.cell(row=row_index + 2, column=col_index+1)
                    cell.fill = red_fill
                    cell.font = font
                    cell.comment = Comment(add_comment(old_value, ref_value), "AutoComparer")
                    ws_summary.append([row_index + 2, df_new_version.columns[col_index], old_value, ref_value])
                    
        wb_new.save(output_path)
        status_label.config(text=f"Comparison complete. Differences highlighted in '{output_path}', summary added in 'Summary' sheet.", fg="green")
    except FileNotFoundError:
        status_label.config(text="Erreur : Fichier(s) introuvable(s).", fg="red")
    except Exception as e:
        status_label.config(text=f"Une erreur s'est produite : {e}", fg="red")

def select_ref_file():
    global ref_file_path
    ref_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    ref_label.config(text=f"Référence: {ref_file_path}")

def select_new_file():
    global new_file_path
    new_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    new_label.config(text=f"Nouveau: {new_file_path}")

def select_output_dir():
    global output_dir
    output_dir = filedialog.askdirectory()
    output_label.config(text=f"Dossier de sortie: {output_dir}")

# Interface Tkinter
root = tk.Tk()
root.title("Excel Comparator - Yazaki")
root.geometry("300x300")

# Ajout du logo
logo_path = "yazaki_logo.png"  # Remplacez par le chemin réel du logo
try:
    img = Image.open(logo_path)
    img = img.resize((150, 50), Image.ANTIALIAS)
    logo = ImageTk.PhotoImage(img)
    logo_label = Label(root, image=logo)
    logo_label.pack(pady=10)
except:
    logo_label = Label(root, text="[Logo Yazaki]", font=("Arial", 14, "bold"))
    logo_label.pack(pady=10)

# Sélection des fichiers
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

# Bouton de comparaison
compare_button = Button(root, text="Comparer", command=compare_excels)
compare_button.pack(pady=20)

# Zone de statut
status_label = Label(root, text="", fg="black")
status_label.pack()

# Exécuter l'interface
tk.mainloop()