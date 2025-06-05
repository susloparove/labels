import fitz

def split_page_into_tiles(
    pdf_path,
    output_path,
    label_size_mm=(75, 30),
    margins_mm=(12, 12, 20, 20),  # left, top, right, bottom
    spacing_mm=(15, 0),            # horizontal, vertical
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

        # Для последней страницы можно изменить число строк
        current_rows = rows if page_num < len(doc) - 1 else 7

        for row in range(current_rows):
            for col in range(cols):
                x0 = left_m + col * (label_w + h_spacing)
                y0 = top_m + row * (label_h + v_spacing)

                clip = fitz.Rect(x0, y0, x0 + label_w, y0 + label_h)

                # Проверка, находится ли область внутри страницы
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


# Пример использования:
split_page_into_tiles(
    "413.pdf",
    "output_labels.pdf",
    rows=9,
    cols=2,
    margins_mm=(12, 12, 20, 20),   # поля в мм
    spacing_mm=(0, 0)              # небольшие отступы между этикетками
)