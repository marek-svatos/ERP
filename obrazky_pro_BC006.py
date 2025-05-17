import os
import zipfile
import pandas as pd
import unicodedata

# Uživatelský vstup
print_success = input("Zobrazovat i úspěšně přidané soubory? (Y/N): ").strip().lower() == "y"
log_missing = input("Zapsat nenalezené soubory do logu 'nenalezeno.txt'? (Y/N): ").strip().lower() == "y"

try:
    max_zip_size_mb = int(input("Maximální velikost jednoho ZIP souboru (v MB, výchozí 100): ").strip())
    if max_zip_size_mb <= 0:
        raise ValueError
except ValueError:
    max_zip_size_mb = 100
    print("→ Použita výchozí velikost: 100 MB")

MAX_ZIP_SIZE = max_zip_size_mb * 1024 * 1024  # převedení na bajty

# Cesty
home = os.path.expanduser("~")
excel_path = os.path.join(home, "Documents", "obrazky.xlsx")
images_root = os.path.join(home, "Documents", "ObrázkyE1")
zip_base_path = os.path.join(home, "Documents", "vybrane_obrazky")
log_path = os.path.join(home, "Documents", "nenalezeno.txt")

# Načtení Excel souboru
df = pd.read_excel(excel_path)

# Získání dvojic: nový název, původní soubor
rows = df.iloc[:, [0, 1]].dropna()
file_pairs = []
for _, row in rows.iterrows():
    new_name = unicodedata.normalize("NFC", str(row.iloc[0]).strip())
    original_filename = unicodedata.normalize("NFC", os.path.basename(str(row.iloc[1]).strip()))
    file_pairs.append((original_filename, new_name))

# Vyhledání všech souborů v adresáři a podsložkách
found_files = {}
for root, dirs, files in os.walk(images_root):
    for fname in files:
        norm_fname = unicodedata.normalize("NFC", fname)
        if norm_fname not in found_files:
            found_files[norm_fname] = os.path.join(root, fname)

# Pomocná funkce pro vytvoření nového ZIP archivu
def create_new_zip(index):
    zip_path = f"{zip_base_path}_{index}.zip"
    return zipfile.ZipFile(zip_path, "w"), zip_path

# Stav
used_output_names = set()
missing_files = []
duplicate_output_names = []
added_files_count = 0

zip_index = 1
current_zip, current_zip_path = create_new_zip(zip_index)
current_size = 0

for original_fname, new_basename in file_pairs:
    file_path = found_files.get(original_fname)
    if file_path:
        _, ext = os.path.splitext(original_fname)
        output_name = f"{new_basename}{ext}"

        if output_name in used_output_names:
            print(f"⚠️  Výstupní název už existuje, soubor přeskočen: {output_name}")
            duplicate_output_names.append(output_name)
            continue

        file_size = os.path.getsize(file_path)
        if current_size + file_size > MAX_ZIP_SIZE:
            current_zip.close()
            print(f"ZIP archiv uložen: {current_zip_path} (celkem {current_size / (1024 * 1024):.2f} MB)")
            zip_index += 1
            current_zip, current_zip_path = create_new_zip(zip_index)
            current_size = 0

        current_zip.write(file_path, arcname=output_name)
        current_size += file_size
        added_files_count += 1
        used_output_names.add(output_name)

        if print_success:
            print(f"Soubor přidán: {original_fname} -> {output_name} ({file_size / (1024 * 1024):.2f} MB)")
    else:
        print(f"Soubor NEnalezen: {original_fname}")
        missing_files.append(original_fname)

# Uzavření posledního ZIPu
current_zip.close()
print(f"ZIP archiv uložen: {current_zip_path} (celkem {current_size / (1024 * 1024):.2f} MB)")

# Logování nenalezených
if log_missing and missing_files:
    with open(log_path, "w", encoding="utf-8") as f:
        for missing in missing_files:
            f.write(missing + "\n")
    print(f"Nenalezené soubory zapsány do: {log_path}")

# Shrnutí
print("\n✅ Hotovo!")
print(f"Přidáno souborů: {added_files_count}")
print(f"Nenalezeno souborů: {len(missing_files)}")
print(f"Přeskočeno duplicitních výstupních názvů: {len(duplicate_output_names)}")
