import os, fitz, subprocess
from fuzzywuzzy import fuzz
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from ttkthemes import ThemedTk


def load_employees_cities(file_path):
    # Load employee and their associated city from file
    employees_cities = {}
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            data = line.strip().split(',')
            employee = data[0].strip()
            city = data[1].strip()
            if len(data) > 2 and city.lower() == 'pinda':
                city = f"{city}_{data[2].strip()}"
            employees_cities[employee] = city
    return employees_cities


def extract_text_from_pdf(pdf_path, region):
    # Extract text from specified region in PDF pages
    pdf_document = fitz.open(pdf_path)
    text_by_page = [page.get_text("text", clip=fitz.Rect(*region)).split("\n") for page in pdf_document]
    pdf_document.close()
    return text_by_page


def find_matching_employee_name(names, text):
    # Find the highest matching employee name in text
    max_ratio = 0
    matched_name = ""
    for name in names:
        for line in text:
            ratio = fuzz.ratio(name.lower(), line.lower())
            if ratio > max_ratio:
                max_ratio, matched_name = ratio, name
    return matched_name if max_ratio >= 95 else None


def separate_pages_by_city(pdf_path, employees_cities, pdfregion):
    # Separate pages in PDF by city based on employee names found
    text_by_page = extract_text_from_pdf(pdf_path, pdfregion)
    pages_by_city = {city: [] for city in set(employees_cities.values())}
    found_names = set()
    for page_num, text in enumerate(text_by_page):
        matched_name = find_matching_employee_name(employees_cities.keys(), text)
        if matched_name:
            found_names.add(matched_name)
            pages_by_city[employees_cities[matched_name]].append(page_num)
    return pages_by_city, set(employees_cities.keys()) - found_names


def save_pages_to_pdf(pdf_path, pages_by_city, output_directory):
    # Save pages sorted by city to individual PDFs
    base_path = os.path.splitext(os.path.basename(pdf_path))[0]
    for city, pages in pages_by_city.items():
        if pages:
            output_pdf_path = os.path.join(output_directory, f'{city.strip().replace(" ", "_")}.pdf')
            with fitz.open(pdf_path) as original_pdf, fitz.open() as pdf_document:
                for page_num in pages:
                    pdf_document.insert_pdf(original_pdf, from_page=page_num, to_page=page_num)
                pdf_document.save(output_pdf_path)


def select_file():
    # File selection dialog
    global selected_pdf
    pdf_file_path = filedialog.askopenfilename(title="Selecionar Arquivo PDF")
    if pdf_file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, pdf_file_path)
        selected_pdf = True


def select_output_directory():
    # Output directory selection dialog
    global selected_output_directory
    output_directory_path = filedialog.askdirectory(title="Selecionar Diretório de Saída")
    if output_directory_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, output_directory_path)
        selected_output_directory = True


def process_pdf():
    # Main processing function for PDF file
    global selected_pdf, selected_output_directory
    if not selected_pdf or not selected_output_directory:
        messagebox.showwarning("Aviso", "Por favor, selecione o arquivo PDF e o diretório de saída.")
        return
    pdf_path, output_directory = file_entry.get(), output_entry.get()
    employees_cities = load_employees_cities("funcionarios_cidade.txt")
    pdfregion = (98, 61, 312, 70)
    pages_by_city, not_found_names = separate_pages_by_city(pdf_path, employees_cities, pdfregion)
    save_pages_to_pdf(pdf_path, pages_by_city, output_directory)
    if not_found_names:
        messagebox.showinfo("Nomes Não Encontrados", "Nomes não encontrados no PDF:\n\n" + "\n".join(f"• {name}" for name in not_found_names))
    messagebox.showinfo("Concluído", "Processo concluído com sucesso!")
    subprocess.Popen(f'explorer "{os.path.abspath(output_directory)}"', shell=True)


selected_pdf = False
selected_output_directory = False

# GUI setup
root = ThemedTk(theme="breeze")
root.title("Doc Separator")
file_label = ttk.Label(root, text="Arquivo PDF:")
file_entry = ttk.Entry(root, width=50)
file_button = ttk.Button(root, text="Selecionar Arquivo", command=select_file)
output_label = ttk.Label(root, text="Diretório de Saída:")
output_entry = ttk.Entry(root, width=50)
output_button = ttk.Button(root, text="Selecionar Diretório", command=select_output_directory)
process_button = ttk.Button(root, text="Processar PDF", command=process_pdf)
file_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
file_entry.grid(row=0, column=1, padx=10, pady=10)
file_button.grid(row=0, column=2, padx=10, pady=10)
output_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
output_entry.grid(row=1, column=1, padx=10, pady=10)
output_button.grid(row=1, column=2, padx=10, pady=10)
process_button.grid(row=2, column=0, columnspan=3, padx=10, pady=10)
root.mainloop()
