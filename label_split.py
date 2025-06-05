import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import fitz  # PyMuPDF
import os


class PDFLabelGridEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Редактор сетки этикеток")

        self.pdf_path = None
        self.tk_img = None
        self.image_scale = 2  # Масштаб рендера для точности
        self.label_size_mm = (58, 40)
        self.margins_mm = (12, 12, 20, 20)
        self.spacing_mm = (2, 2)
        self.rows = 9
        self.cols = 2

        # Создание интерфейса
        self.create_widgets()
        self.select_file()

    def create_widgets(self):
        # Canvas с прокруткой
        self.canvas_frame = ttk.Frame(self.root)
        self.canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.canvas_frame)
        self.scroll_x = tk.Scrollbar(self.canvas_frame, orient="horizontal", command=self.canvas.xview)
        self.scroll_y = tk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self.scroll_x.set, yscrollcommand=self.scroll_y.set)

        self.scroll_x.pack(side="bottom", fill="x")
        self.scroll_y.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Панель управления
        control_frame = ttk.Frame(self.root)
        control_frame.pack(pady=10, fill="x")

        ttk.Label(control_frame, text="Размер этикетки (мм):").grid(row=0, column=0)
        self.entry_w = ttk.Entry(control_frame, width=6)
        self.entry_h = ttk.Entry(control_frame, width=6)
        self.entry_w.insert(0, str(self.label_size_mm[0]))
        self.entry_h.insert(0, str(self.label_size_mm[1]))
        self.entry_w.grid(row=0, column=1)
        self.entry_h.grid(row=0, column=2)

        ttk.Label(control_frame, text="Поля (л,в,п,н):").grid(row=0, column=3)
        self.entry_ml = ttk.Entry(control_frame, width=5)
        self.entry_mt = ttk.Entry(control_frame, width=5)
        self.entry_mr = ttk.Entry(control_frame, width=5)
        self.entry_mb = ttk.Entry(control_frame, width=5)
        for i, e in enumerate([self.entry_ml, self.entry_mt, self.entry_mr, self.entry_mb]):
            e.insert(0, str(self.margins_mm[i]))
            e.grid(row=0, column=4 + i)

        ttk.Label(control_frame, text="Зазоры (мм):").grid(row=0, column=8)
        self.entry_sx = ttk.Entry(control_frame, width=5)
        self.entry_sy = ttk.Entry(control_frame, width=5)
        self.entry_sx.insert(0, str(self.spacing_mm[0]))
        self.entry_sy.insert(0, str(self.spacing_mm[1]))
        self.entry_sx.grid(row=0, column=9)
        self.entry_sy.grid(row=0, column=10)

        ttk.Label(control_frame, text="Строк/Столбцов:").grid(row=0, column=11)
        self.entry_rows = ttk.Entry(control_frame, width=5)
        self.entry_cols = ttk.Entry(control_frame, width=5)
        self.entry_rows.insert(0, str(self.rows))
        self.entry_cols.insert(0, str(self.cols))
        self.entry_rows.grid(row=0, column=12)
        self.entry_cols.grid(row=0, column=13)

        ttk.Button(control_frame, text="Обновить", command=self.update_grid).grid(row=0, column=14, padx=5)
        ttk.Button(control_frame, text="Применить", command=self.apply_settings).grid(row=0, column=15, padx=5)

    def select_file(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")])
        if not self.pdf_path:
            messagebox.showerror("Ошибка", "Файл не выбран")
            return
        self.load_page_with_grid()

    def load_page_with_grid(self):
        doc = fitz.open(self.pdf_path)
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(self.image_scale, self.image_scale))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.tk_img = ImageTk.PhotoImage(img)
        doc.close()

        self.image_on_canvas = self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        self.canvas.config(scrollregion=(0, 0, self.tk_img.width(), self.tk_img.height()))

        self.draw_grid()

    def mm_to_px(self, mm):
        return mm * 2.8346 * self.image_scale

    def draw_grid(self):
        self.canvas.delete("grid")

        label_w = self.mm_to_px(float(self.entry_w.get()))
        label_h = self.mm_to_px(float(self.entry_h.get()))
        spacing_x = self.mm_to_px(float(self.entry_sx.get()))
        spacing_y = self.mm_to_px(float(self.entry_sy.get()))
        margin_l = self.mm_to_px(float(self.entry_ml.get()))
        margin_t = self.mm_to_px(float(self.entry_mt.get()))
        rows = int(self.entry_rows.get())
        cols = int(self.entry_cols.get())

        for row in range(rows):
            for col in range(cols):
                x0 = margin_l + col * (label_w + spacing_x)
                y0 = margin_t + row * (label_h + spacing_y)
                self.canvas.create_rectangle(
                    x0, y0,
                    x0 + label_w, y0 + label_h,
                    outline="red",
                    dash=(4, 4),
                    tags="grid"
                )

    def update_grid(self):
        self.canvas.delete("grid")
        self.draw_grid()

    def apply_settings(self):
        try:
            self.label_size_mm = (float(self.entry_w.get()), float(self.entry_h.get()))
            self.margins_mm = tuple(float(e.get()) for e in [self.entry_ml, self.entry_mt, self.entry_mr, self.entry_mb])
            self.spacing_mm = (float(self.entry_sx.get()), float(self.entry_sy.get()))
            self.rows = int(self.entry_rows.get())
            self.cols = int(self.entry_cols.get())

            # Здесь можно вызвать функцию разрезания PDF
            split_page_into_tiles(
                self.pdf_path,
                "output_labels.pdf",
                label_size_mm=self.label_size_mm,
                margins_mm=self.margins_mm,
                spacing_mm=self.spacing_mm,
                rows=self.rows,
                cols=self.cols
            )
            messagebox.showinfo("Готово", "Файл успешно разделён!")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


def split_page_into_tiles(
    pdf_path,
    output_path,
    label_size_mm=(58, 40),
    margins_mm=(12, 12, 20, 20),
    spacing_mm=(2, 2),
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
    h_spacing = spacing_mm[0] * mm_to_point
    v_spacing = spacing_mm[1] * mm_to_point

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        page_rect = page.rect

        current_rows = rows if page_num < len(doc) - 1 else 7

        for row in range(current_rows):
            for col in range(cols):
                x0 = left_m + col * (label_w + h_spacing)
                y0 = top_m + row * (label_h + v_spacing)
                clip = fitz.Rect(x0, y0, x0 + label_w, y0 + label_h)

                if not clip.is_valid or clip.is_empty:
                    continue
                if clip.x1 > page_rect.width or clip.y1 > page_rect.height:
                    continue

                new_page = new_doc.new_page(width=label_w, height=label_h)
                new_page.show_pdf_page(new_page.rect, doc, page_num, clip=clip)

    new_doc.save(output_path)
    new_doc.close()
    doc.close()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFLabelGridEditor(root)
    root.mainloop()