import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import os


def split_page_into_tiles(
    pdf_path,
    output_path,
    label_size_mm=(75, 30),
    margins_mm=(12, 12, 20, 20),
    spacing_mm=(15, 0),
    rows=9,
    cols=2
):
    doc = fitz.open(pdf_path)
    new_doc = fitz.open()

    mm_to_point = 2.8346
    label_w = label_size_mm[0] * mm_to_point
    label_h = label_size_mm[1] * mm_to_point

    left_m = margins_mm[0] * mm_to_point
    top_m = margins_mm[1] * mm_to_point
    right_m = margins_mm[2] * mm_to_point
    bottom_m = margins_mm[3] * mm_to_point

    h_spacing = spacing_mm[0] * mm_to_point
    v_spacing = spacing_mm[1] * mm_to_point

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        page_rect = page.rect

        print(f"\nСтраница {page_num}, размер: {page_rect}")

        current_rows = rows if page_num < len(doc) - 1 else 7  # На последней странице меньше строк

        for row in range(current_rows):
            for col in range(cols):
                x0 = left_m + col * (label_w + h_spacing)
                y0 = top_m + row * (label_h + v_spacing)

                clip = fitz.Rect(x0, y0, x0 + label_w, y0 + label_h)

                if not clip.is_valid or clip.is_empty:
                    print(f"Пропущена область: ({x0}, {y0}) - некорректная")
                    continue
                if clip.x1 > page_rect.width or clip.y1 > page_rect.height:
                    print(f"Пропущена область: {clip} - выходит за границы страницы")
                    continue

                print(f"Обрабатываем область: {clip}")
                new_page = new_doc.new_page(width=label_w, height=label_h)
                new_page.show_pdf_page(new_page.rect, doc, page_num, clip=clip)

    new_doc.save(output_path)
    new_doc.close()
    doc.close()
    print(f"\nФайл успешно сохранён как {output_path}")


class PDFLabelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Разделение этикеток из PDF")

        # Поля ввода
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="Файл PDF:").grid(row=0, column=0, sticky="w")
        self.file_path = tk.Entry(self.root, width=50)
        self.file_path.grid(row=0, column=1)
        tk.Button(self.root, text="Выбрать", command=self.select_file).grid(row=0, column=2)

        # Размер этикетки
        tk.Label(self.root, text="Размер этикетки (мм):").grid(row=1, column=0, sticky="w")
        self.label_w = tk.Entry(self.root, width=10)
        self.label_h = tk.Entry(self.root, width=10)
        self.label_w.insert(0, "75")
        self.label_h.insert(0, "30")
        self.label_w.grid(row=1, column=1, sticky="w")
        self.label_h.grid(row=1, column=1)

        # Строки и столбцы
        tk.Label(self.root, text="Число строк и столбцов:").grid(row=2, column=0, sticky="w")
        self.rows = tk.Entry(self.root, width=10)
        self.cols = tk.Entry(self.root, width=10)
        self.rows.insert(0, "9")
        self.cols.insert(0, "2")
        self.rows.grid(row=2, column=1, sticky="w")
        self.cols.grid(row=2, column=1)

        # Поля
        tk.Label(self.root, text="Поля (лево, верх, право, низ) мм:").grid(row=3, column=0, sticky="w")
        self.margin_l = tk.Entry(self.root, width=5)
        self.margin_t = tk.Entry(self.root, width=5)
        self.margin_r = tk.Entry(self.root, width=5)
        self.margin_b = tk.Entry(self.root, width=5)

        # Устанавливаем значения по умолчанию
        self.margin_l.insert(0, "12")  # лево
        self.margin_t.insert(0, "12")  # верх
        self.margin_r.insert(0, "20")  # право
        self.margin_b.insert(0, "20")  # низ

        # Размещаем поля ввода рядом друг с другом
        self.margin_l.grid(row=3, column=1, padx=(0, 10))
        self.margin_t.grid(row=3, column=1, padx=(60, 10))
        self.margin_r.grid(row=3, column=1, padx=(120, 10))
        self.margin_b.grid(row=3, column=1, padx=(180, 0))

        # Отступы между ячейками
        tk.Label(self.root, text="Отступы между этикетками (гор., верт.) мм:").grid(row=4, column=0, sticky="w")
        self.spacing_x = tk.Entry(self.root, width=10)
        self.spacing_y = tk.Entry(self.root, width=10)
        self.spacing_x.insert(0, "15")
        self.spacing_y.insert(0, "0")
        self.spacing_x.grid(row=4, column=1, sticky="w")
        self.spacing_y.grid(row=4, column=1)

        # Кнопка запуска
        tk.Button(self.root, text="Запустить", command=self.run).grid(row=5, column=1, pady=10)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")])
        if path:
            self.file_path.delete(0, tk.END)
            self.file_path.insert(0, path)

    def run(self):
        try:
            pdf_path = self.file_path.get()
            if not os.path.isfile(pdf_path):
                raise FileNotFoundError("Укажите корректный путь к PDF-файлу")

            label_size_mm = (float(self.label_w.get()), float(self.label_h.get()))
            rows = int(self.rows.get())
            cols = int(self.cols.get())
            margins_mm = (
                float(self.margin_l.get()),
                float(self.margin_t.get()),
                float(self.margin_r.get()),
                float(self.margin_b.get())
            )
            spacing_mm = (float(self.spacing_x.get()), float(self.spacing_y.get()))

            output_path = os.path.splitext(pdf_path)[0] + "_split.pdf"

            split_page_into_tiles(
                pdf_path,
                output_path,
                label_size_mm=label_size_mm,
                margins_mm=margins_mm,
                spacing_mm=spacing_mm,
                rows=rows,
                cols=cols
            )

            messagebox.showinfo("Готово", f"Файл сохранён как:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFLabelSplitterApp(root)
    root.mainloop()