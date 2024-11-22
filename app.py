import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from ttkthemes import ThemedTk


def replace_all(text, dic):
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
    global selected_pdf
    pdf_file_path = filedialog.askopenfilename(title="Selecionar Arquivo PDF")
    if pdf_file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, pdf_file_path)
        selected_pdf = True
        settings_button['state'] = 'normal'
        open_settings()


def select_output_directory():
    # Output directory selection dialog
    global selected_output_directory
    output_directory_path = filedialog.askdirectory(title="Selecionar Diretório de Saída")
    if output_directory_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, output_directory_path)
        selected_output_directory = True


def process_pdf():
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
    pdfregion = load_pdfregion()  # Carregar coordenadas atualizadas
    pages_by_city, nfn_by_city = separate_pages_by_city(pdf_path, employees_cities, pdfregion)
    save_pages_to_pdf(pdf_path, pages_by_city, output_directory)
    root.config(cursor="")  # Resetar o cursor aqui

    if nfn_by_city:
        display_not_found_names(nfn_by_city)
    messagebox.showinfo("Concluído", "Processo concluído com sucesso!")
    Popen(f'explorer "{os.path.abspath(output_directory)}"', shell=True)


def display_not_found_names(nfn_by_city):
    # Criar uma nova janela Toplevel
    nfn_window = tk.Toplevel(root)
    nfn_window.title("Nomes Não Encontrados")

    # Obter as dimensões da tela
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Definir o tamanho máximo da janela (por exemplo, 50% da tela)
    max_window_width = int(screen_width * 0.5)
    max_window_height = int(screen_height * 0.5)

    # Definir o tamanho da janela
    window_width = min(400, max_window_width)
    window_height = min(300, max_window_height)

    # Centralizar a janela na tela
    x_position = int((screen_width / 2) - (window_width / 2))
    y_position = int((screen_height / 2) - (window_height / 2))
    nfn_window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    # Tornar a janela modal
    nfn_window.transient(root)
    nfn_window.grab_set()
    nfn_window.focus()

    # Impedir redimensionamento da janela (opcional)
    nfn_window.resizable(False, False)

    # Criar um Frame para conter o Text widget e a barra de rolagem
    frame = tk.Frame(nfn_window)
    frame.pack(fill=tk.BOTH, expand=True)

    # Criar a barra de rolagem vertical
    v_scroll = tk.Scrollbar(frame, orient=tk.VERTICAL)
    v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # Criar o widget de texto
    text_widget = tk.Text(frame, wrap=tk.WORD, yscrollcommand=v_scroll.set)
    text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Configurar a barra de rolagem
    v_scroll.config(command=text_widget.yview)

    # Construir o conteúdo da mensagem
    nfn_msg = []
    for city, names in nfn_by_city.items():
        if names:
            city_section = f"{city}\n"
            city_section += "\n".join(f"• {name}" for name in names)
            city_section += "\n\n"
            nfn_msg.append(city_section)

    message_content = "Estes nomes não foram encontrados no PDF:\n\n" + "".join(nfn_msg)

    # Inserir o conteúdo no widget de texto
    text_widget.insert(tk.END, message_content)

    # Desabilitar a edição do texto
    text_widget.config(state=tk.DISABLED)

    # Adicionar botão "Ok" para fechar a janela
    ok_button = ttk.Button(nfn_window, text="Ok", command=nfn_window.destroy)
    ok_button.pack(pady=5)

    # Aguardar até que a janela seja fechada
    nfn_window.wait_window()


def load_pdfregion():
    try:
        with open('pdfregion.txt', 'r') as f:
            coords = f.read().split(',')
            return list(map(int, coords))
    except FileNotFoundError:
        return [100, 100, 300, 200]


def save_pdfregion_and_close(canvas, rect, settings_window, zoom=1.0):
    x1, y1, x2, y2 = map(int, canvas.coords(rect))
    # Ajustar coordenadas de volta ao tamanho original
    x1, y1, x2, y2 = [int(coord / zoom) for coord in (x1, y1, x2, y2)]
    with open('pdfregion.txt', 'w') as f:
        f.write(f"{x1},{y1},{x2},{y2}")
    settings_window.destroy()
    messagebox.showinfo("Configurações Salvas", "As coordenadas foram salvas com sucesso.")


def open_settings():
    settings_window = tk.Toplevel(root)
    settings_window.title("Configurações")

    pdf_path = file_entry.get()
    import fitz
    pdf_document = fitz.open(pdf_path)
    page = pdf_document.load_page(0)
    zoom = 2.0  # Zoom de 200%
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    pdf_document.close()

    from PIL import Image, ImageTk
    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    photo = ImageTk.PhotoImage(image)

    # Obter a largura e altura da imagem
    image_width, image_height = image.size

    # Verificar a largura da tela
    screen_width = root.winfo_screenwidth()
    if image_width > screen_width:
        window_width = screen_width
        h_scroll_needed = True
    else:
        window_width = image_width
        h_scroll_needed = False

    window_height = 600  # Ajuste conforme necessário
    settings_window.geometry(f"{window_width}x{window_height}")

    # Criar um Frame para o Canvas com barras de rolagem
    canvas_frame = tk.Frame(settings_window)
    canvas_frame.pack(fill=tk.BOTH, expand=True)

    # Criar a barra de rolagem vertical
    v_scroll = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
    v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # Criar o Canvas
    canvas = tk.Canvas(canvas_frame, width=window_width,
                       height=window_height - 50,  # Espaço para os botões
                       yscrollcommand=v_scroll.set)

    # Configurar a barra de rolagem vertical
    v_scroll.config(command=canvas.yview)

    # Criar e configurar a barra de rolagem horizontal se necessário
    if h_scroll_needed:
        h_scroll = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.config(xscrollcommand=h_scroll.set)
        h_scroll.config(command=canvas.xview)

    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Adicionar a imagem ao Canvas
    canvas.create_image(0, 0, anchor='nw', image=photo)
    canvas.image = photo

    # Definir a região de rolagem do Canvas
    canvas.config(scrollregion=(0, 0, image_width, image_height))

    pdfregion = load_pdfregion()
    # Ajustar coordenadas de acordo com o zoom
    x1, y1, x2, y2 = [coord * zoom for coord in pdfregion]

    # Desenhar o retângulo com preenchimento semi-transparente
    rect = canvas.create_rectangle(x1, y1, x2, y2, outline='red',
                                   fill='gray', stipple='gray25', width=2)

    # Criar os handles para redimensionamento
    handle_size = 6
    half_handle = handle_size // 2

    handles = {}
    handles['nw'] = canvas.create_rectangle(x1 - half_handle, y1 - half_handle,
                                            x1 + half_handle, y1 + half_handle, fill='blue')
    handles['ne'] = canvas.create_rectangle(x2 - half_handle, y1 - half_handle,
                                            x2 + half_handle, y1 + half_handle, fill='blue')
    handles['sw'] = canvas.create_rectangle(x1 - half_handle, y2 - half_handle,
                                            x1 + half_handle, y2 + half_handle, fill='blue')
    handles['se'] = canvas.create_rectangle(x2 - half_handle, y2 - half_handle,
                                            x2 + half_handle, y2 + half_handle, fill='blue')

    # Garantir que os handles estejam acima do retângulo
    for handle in handles.values():
        canvas.tag_raise(handle)

    rr = ResizableRectangle(canvas, rect, handles, zoom=zoom)

    # Botões Confirmar e Cancelar
    button_frame = tk.Frame(settings_window)
    button_frame.pack(pady=5)

    confirm_button = ttk.Button(button_frame, text="Confirmar",
                                command=lambda: save_pdfregion_and_close(canvas, rect, settings_window, zoom))
    confirm_button.pack(side='left', padx=10)

    cancel_button = ttk.Button(button_frame, text="Cancelar", command=settings_window.destroy)
    cancel_button.pack(side='right', padx=10)


class ResizableRectangle:
    def __init__(self, canvas, rect, handles, zoom=1.0):
        self.canvas = canvas
        self.rect = rect
        self.handles = handles
        self.zoom = zoom
        self._drag_data = {"x": 0, "y": 0, "item": None}

        # Eventos para mover o retângulo
        canvas.tag_bind(rect, "<ButtonPress-1>", self.on_rect_button_press)
        canvas.tag_bind(rect, "<ButtonRelease-1>", self.on_rect_button_release)
        canvas.tag_bind(rect, "<B1-Motion>", self.on_rect_motion)

        # Eventos para redimensionar usando os handles
        for handle_name, handle in handles.items():
            canvas.tag_bind(handle, "<ButtonPress-1>", self.on_handle_button_press)
            canvas.tag_bind(handle, "<ButtonRelease-1>", self.on_handle_button_release)
            canvas.tag_bind(handle, "<B1-Motion>", self.on_handle_motion)

    def canvas_coords(self, event):
        return self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

    def on_rect_button_press(self, event):
        x, y = self.canvas_coords(event)
        self._drag_data["item"] = self.rect
        self._drag_data["x"] = x
        self._drag_data["y"] = y

    def on_rect_button_release(self, event):
        self._drag_data["item"] = None
        self._drag_data["x"] = 0
        self._drag_data["y"] = 0

    def on_rect_motion(self, event):
        x, y = self.canvas_coords(event)
        dx = x - self._drag_data["x"]
        dy = y - self._drag_data["y"]
        self.canvas.move(self.rect, dx, dy)
        for handle in self.handles.values():
            self.canvas.move(handle, dx, dy)
        self._drag_data["x"] = x
        self._drag_data["y"] = y

    def on_handle_button_press(self, event):
        x, y = self.canvas_coords(event)
        self._drag_data["item"] = self.canvas.find_withtag("current")[0]
        self._drag_data["x"] = x
        self._drag_data["y"] = y

    def on_handle_button_release(self, event):
        self._drag_data["item"] = None
        self._drag_data["x"] = 0
        self._drag_data["y"] = 0

    def on_handle_motion(self, event):
        x, y = self.canvas_coords(event)
        dx = x - self._drag_data["x"]
        dy = y - self._drag_data["y"]
        handle = self._drag_data["item"]
        x1, y1, x2, y2 = self.canvas.coords(self.rect)

        if handle == self.handles['nw']:
            x1 += dx
            y1 += dy
        elif handle == self.handles['ne']:
            x2 += dx
            y1 += dy
        elif handle == self.handles['sw']:
            x1 += dx
            y2 += dy
        elif handle == self.handles['se']:
            x2 += dx
            y2 += dy

        self.canvas.coords(self.rect, x1, y1, x2, y2)
        self.update_handles()

        self._drag_data["x"] = x
        self._drag_data["y"] = y

    def update_handles(self):
        x1, y1, x2, y2 = self.canvas.coords(self.rect)
        half_handle = 3
        self.canvas.coords(self.handles['nw'], x1 - half_handle, y1 - half_handle,
                           x1 + half_handle, y1 + half_handle)
        self.canvas.coords(self.handles['ne'], x2 - half_handle, y1 - half_handle,
                           x2 + half_handle, y1 + half_handle)
        self.canvas.coords(self.handles['sw'], x1 - half_handle, y2 - half_handle,
                           x1 + half_handle, y2 + half_handle)
        self.canvas.coords(self.handles['se'], x2 - half_handle, y2 - half_handle,
                           x2 + half_handle, y2 + half_handle)


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

# Adicionar botão de configurações
settings_button = ttk.Button(root, text="Configurações", command=open_settings)
settings_button.grid(row=0, column=3, padx=10, pady=10)
settings_button['state'] = 'disabled'

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
