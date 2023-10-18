import tkinter as tk
from tkinter import filedialog
import openpyxl
import random

existing_codes = set()  # Déclarer la variable existing_codes en dehors des fonctions

def extract_prefixes(existing_codes):
    prefixes = set()
    for code in existing_codes:
        prefix = code[:6]  # Extraction des 6 premiers chiffres du code EAN-13
        prefixes.add(prefix)
    return list(prefixes)

def generate_random_ean13(prefixes, num_codes):
    new_codes = []
    while len(new_codes) < num_codes:
        prefix = random.choice(prefixes)  # Choix aléatoire d'un préfixe existant
        random_digits = ''.join([str(random.randint(0, 9)) for _ in range(6)])  # 6 chiffres aléatoires
        ean13 = prefix + random_digits
        checksum = 0
        for i, digit in enumerate(ean13):
            digit = int(digit)
            if i % 2 == 0:
                checksum += digit
            else:
                checksum += digit * 3
        checksum = (10 - (checksum % 10)) % 10
        ean13 += str(checksum)
        if ean13 not in existing_codes and ean13 not in new_codes:
            new_codes.append(ean13)
    return new_codes

def process_excel_file(num_codes):
    file_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx")])
    
    if file_path:
        try:
            existing_wb = openpyxl.load_workbook(file_path)
            existing_sheet = existing_wb.active
            existing_codes.clear()  # Effacer les codes existants pour la mise à jour

            # Parcourt les cellules du fichier Excel et récupère les valeurs
            for row in existing_sheet.iter_rows(values_only=True):
                for cell in row:
                    cell_str = str(cell)
                    if len(cell_str) == 13 and cell_str.isdigit():
                        existing_codes.add(cell_str)

            # Extraire les préfixes des codes EAN-13 existants
            prefixes = extract_prefixes(existing_codes)

            # Générer les nouveaux codes EAN-13 avec les préfixes existants
            new_codes = generate_random_ean13(prefixes, num_codes)

            # Ajouter les nouveaux codes au fichier Excel
            for code in new_codes:
                existing_sheet.append([code])

            new_file_path = file_path.replace('.xlsx', '_modified.xlsx')

            # Sauvegarder le fichier Excel modifié
            existing_wb.save(new_file_path)
            existing_wb.close()
            
            status_label.config(text=f"{len(new_codes)} codes EAN-13 ont été générés et ajoutés dans le fichier Excel modifié : {new_file_path}")
        except Exception as e:
            status_label.config(text=f"Erreur : {str(e)}")

root = tk.Tk()
root.title("Générateur de codes EAN-13 dans un fichier Excel")

select_file_button = tk.Button(root, text="Sélectionner un fichier Excel", command=lambda: process_excel_file(int(entry.get())))
select_file_button.pack(pady=20)

num_codes_label = tk.Label(root, text="Nombre de codes à générer :")
num_codes_label.pack()
entry = tk.Entry(root)
entry.pack()

status_label = tk.Label(root, text="")
status_label.pack()

root.mainloop()
