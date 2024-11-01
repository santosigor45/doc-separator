import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from ttkthemes import ThemedTk


def replace_all(text, dic):
    # A function to replace all passed characters
    for i, j in dic.items():
        text = text.replace(i, j)
    return text


def load_employees_cities(file_path):
    # Load employee and their associated city from file
    from numpy import nan
    from pandas import read_excel

    employees_cities = {}
    excel_file = read_excel(
        io=file_path,
        sheet_name="GERAL",
        usecols=["FUNCIONÁRIO", "CIDADE", "SITUAÇÃO", "REGIÃO"],
    )
    excel_file = excel_file[(excel_file['SITUAÇÃO'] == 'REGISTRADO')]
    excel_file = excel_file.drop('SITUAÇÃO', axis=1)
    excel_file = excel_file.replace(nan, '', regex=True)

    lists = excel_file.values.tolist()

    for ls in lists:
        employee = ls[0].strip()
        city = ls[1].strip()
        if len(ls) > 2 and ls[2] != '' and regions_checkbox.state() == ('selected',):
            region = replace_all(ls[2], {' ': '_', '/': ', '}).strip()
            city = f"{city} - {region}"
        employees_cities[employee] = city
    return employees_cities


def extract_text_from_pdf(pdf_path, region):
    # Extract text from specified region in PDF pages
    from fitz import open as fitz_open, Rect
    pdf_document = fitz_open(pdf_path)
    text_by_page = [page.get_text("text", clip=Rect(*region)).split("\n")[0] for page in pdf_document]
    pdf_document.close()
    return text_by_page


def find_matching_employee_name(names, text):
    # Find the highest matching employee name in text
    from fuzzywuzzy import fuzz
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
        matched_name = find_matching_employee_name(employees_cities.keys(), text.split("\n"))
        if matched_name:
            found_names.add(matched_name)
            pages_by_city[employees_cities[matched_name]].append(page_num)

    not_found_names = sorted(set(employees_cities.keys()) - found_names)
    only_cities = []
    for city in set(employees_cities.values()):
        only_cities.append(city.split(' - ')[0])
    nfn_by_city = {city: [] for city in sorted(set(only_cities))}
    for city in nfn_by_city.keys():
        for name in not_found_names:
            if city == employees_cities[name].split(' - ')[0]:
                nfn_by_city[city].append(name)

    return pages_by_city, nfn_by_city


def save_pages_to_pdf(pdf_path, pages_by_city, output_directory):
    # Save pages sorted by city to individual PDFs
    from fitz import open as fitz_open
    for city, pages in pages_by_city.items():
        if pages:
            output_pdf_path = os.path.join(output_directory, f'{city.strip()}.pdf')
            with fitz_open(pdf_path) as original_pdf, fitz_open() as pdf_document:
                for page_num in pages:
                    pdf_document.insert_pdf(original_pdf, from_page=page_num, to_page=page_num)

                if len(city.split(' - ')) == 2:
                    city_name = city.split(' - ')[0]
                    new_folder = os.path.join(output_directory, city_name)
                    if not os.path.exists(new_folder):
                        os.makedirs(new_folder)
                    pdf_output_path = os.path.join(new_folder, f'{city.strip().replace(f"{city_name} - ", "")}.pdf')
                    pdf_document.save(pdf_output_path)
                else:
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
    from pathlib import Path
    from subprocess import Popen
    global selected_pdf, selected_output_directory
    root.config(cursor="watch")
    root.update()
    if not selected_pdf or not selected_output_directory:
        messagebox.showwarning("Aviso", "Por favor, selecione o arquivo PDF e o diretório de saída.")
        return
    pdf_path, output_directory = file_entry.get(), output_entry.get()
    employees_cities = load_employees_cities(os.path.abspath(Path('excel_path').read_bytes()).decode('utf-8'))
    # pdfregion = (98, 61, 312, 70)  # old payslip size
    pdfregion = (84, 48, 360, 55)  # new payslip size
    pages_by_city, nfn_by_city = separate_pages_by_city(pdf_path, employees_cities, pdfregion)
    save_pages_to_pdf(pdf_path, pages_by_city, output_directory)
    if nfn_by_city:
        nfn_msg = []
        i = 0
        for city in nfn_by_city.keys():
            if len(nfn_by_city[city]) > 0:
                nfn_msg.append(f"{city}\n")
                for name in nfn_by_city[city]:
                    nfn_msg[i] += f"• {name}\n"
            else:
                i -= 1
            i += 1
        messagebox.showinfo("Nomes Não Encontrados", "Estes nomes não foram encontrados no PDF:\n\n" + "\n"
                            .join(f"{city}" for city in nfn_msg))
    root.config(cursor="")
    messagebox.showinfo("Concluído", "Processo concluído com sucesso!")
    Popen(f'explorer "{os.path.abspath(output_directory)}"', shell=True)


selected_pdf = False
selected_output_directory = False

# GUI setup
root = ThemedTk(theme="breeze")
root.title("Doc Separator")

file_label = ttk.Label(root, text="Arquivo PDF:")
file_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

file_entry = ttk.Entry(root, width=50)
file_entry.grid(row=0, column=1, padx=10, pady=10)

file_button = ttk.Button(root, text="Selecionar Arquivo", command=select_file)
file_button.grid(row=0, column=2, padx=10, pady=10)

output_label = ttk.Label(root, text="Diretório de Saída:")
output_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

output_entry = ttk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)

output_button = ttk.Button(root, text="Selecionar Diretório", command=select_output_directory)
output_button.grid(row=1, column=2, padx=10, pady=10)

regions_checkbox = ttk.Checkbutton(root, text="Separar por regiões?")
regions_checkbox.grid(row=2, column=2, columnspan=1, padx=10, pady=10)

process_button = ttk.Button(root, text="Processar PDF", command=process_pdf)
process_button.grid(row=2, column=0, columnspan=3, padx=10, pady=10)

root.mainloop()
