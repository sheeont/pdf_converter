from tkinter import filedialog, Tk, Button, Label, messagebox, Toplevel, Listbox, EXTENDED, END, Canvas, Scrollbar, Frame
from PIL import Image, ImageTk
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF


def images_to_pdf(image_paths, output_pdf_path):
    """Конвертирует изображения в PDF."""
    images = [Image.open(img_path).convert('RGB') for img_path in image_paths]
    images[0].save(output_pdf_path, save_all=True, append_images=images[1:])
    messagebox.showinfo("Успех", f"PDF создан: {output_pdf_path}")


def select_images():
    """Выбор изображений для конвертации в PDF."""
    image_paths = filedialog.askopenfilenames(title="Выберите изображения", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")])
    if image_paths:
        output_pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_pdf_path:
            images_to_pdf(image_paths, output_pdf_path)


def split_pdf(input_pdf_path, page_indices, output_pdf_path):
    """Создает новый PDF, содержащий только указанные страницы."""
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    for index in page_indices:
        writer.add_page(reader.pages[index])

    with open(output_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)
    messagebox.showinfo("Успех", f"PDF разделен: {output_pdf_path}")


def merge_pdfs(selected_pages_order):
    """Создает новый PDF из выбранных страниц в заданном порядке."""
    writer = PdfWriter()

    for pdf_path, page_num in selected_pages_order:
        reader = PdfReader(pdf_path)
        writer.add_page(reader.pages[page_num])

    output_pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    if output_pdf_path:
        with open(output_pdf_path, "wb") as output_pdf:
            writer.write(output_pdf)
        messagebox.showinfo("Успех", f"PDF успешно объединен: {output_pdf_path}")


def preview_pdf_pages(pdf_path, selected_pages_order):
    """Открывает новое окно для предпросмотра страниц PDF и добавления их в итоговое объединение."""
    doc = fitz.open(pdf_path)
    preview_window = Toplevel(root)
    preview_window.title(f"Предпросмотр страниц {pdf_path.split('/')[-1]}")

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    preview_window.geometry(f"{int(screen_width * 0.4)}x{int(screen_height * 0.4)}")
    preview_window.minsize(600, 400)

    canvas = Canvas(preview_window)
    scrollbar = Scrollbar(preview_window, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas)

    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    def on_page_click(event, page_num):
        """Обработчик клика по изображению страницы. Добавляет или убирает страницу из списка для объединения."""
        if (pdf_path, page_num) in selected_pages_order:
            selected_pages_order.remove((pdf_path, page_num))
            event.widget.config(borderwidth=0, relief="flat")
        else:
            selected_pages_order.append((pdf_path, page_num))
            event.widget.config(borderwidth=2, relief="solid")

    thumbnail_width = 250
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=75)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((thumbnail_width, int(thumbnail_width * img.height / img.width)))
        img_tk = ImageTk.PhotoImage(img)

        label = Label(scrollable_frame, image=img_tk)
        label.image = img_tk
        label.grid(row=i // 2, column=i % 2, padx=5, pady=5)
        label.bind("<Button-1>", lambda event, i=i: on_page_click(event, i))

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    Button(preview_window, text="Закрыть", command=preview_window.destroy).pack(pady=10)


def select_pdfs_to_merge():
    """Выбор PDF файлов для объединения и создание общего PDF."""
    pdf_paths = filedialog.askopenfilenames(title="Выберите PDF файлы для объединения", filetypes=[("PDF Files", "*.pdf")])
    if pdf_paths:
        selected_pages_order = []

        for pdf_path in pdf_paths:
            preview_pdf_pages(pdf_path, selected_pages_order)

        def save_merged_pdf():
            if selected_pages_order:
                merge_pdfs(selected_pages_order)

        # Окно сохранения объединённого PDF
        merge_window = Toplevel(root)
        merge_window.title("Сохранить объединённый PDF")
        Button(merge_window, text="Создать объединённый PDF", command=save_merged_pdf).pack(pady=20)


def select_pdf_to_split():
    """Выбор PDF файла для разделения."""
    input_pdf_path = filedialog.askopenfilename(title="Выберите PDF файл", filetypes=[("PDF Files", "*.pdf")])
    if input_pdf_path:
        selected_pages_order = []
        preview_pdf_pages(input_pdf_path, selected_pages_order)


# Настройка графического интерфейса
root = Tk()
root.title("PDF Manipulator")

Label(root, text="Выберите действие:").pack(pady=10)

Button(root, text="Конвертировать изображения в PDF", command=select_images).pack(pady=5)
Button(root, text="Разделить PDF", command=select_pdf_to_split).pack(pady=5)
Button(root, text="Объединить PDF файлы", command=select_pdfs_to_merge).pack(pady=5)

root.mainloop()
