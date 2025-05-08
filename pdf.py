import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import datetime

excel_file_input = "checklist_sample.xlsx"  # change for your file path
base_title = "Recipe Preparation Checklist for Culinary Approval"  # change

#
#


def create_interactive_pdf_v4(excel_path, pdf_title_input):
    """
    Creates an interactive PDF from an Excel file with improved layout:
    - Adds a wide, single-line edit field below the main title (first page only).
    - Checkbox column is very narrow (fixed width).
    - Edit fields in the table have a minimum height of ~3 lines.
    - Row height is dynamic based on content.
    - Corrected form object scoping.
    """
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook.active
    except FileNotFoundError:
        print(f"Erro: Arquivo Excel não encontrado em '{excel_path}\"")
        return None
    except Exception as e:
        print(f"Erro ao abrir o arquivo Excel: {e}")
        return None

    main_font = "Helvetica"

    pdf_file_name = f"{pdf_title_input.replace(' ', '_')}.pdf"
    c = canvas.Canvas(pdf_file_name, pagesize=A4)
    width, height = A4

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["h1"],
        fontName=main_font,
        alignment=TA_CENTER,
        fontSize=16,
        leading=20,
    )
    header_style = ParagraphStyle(
        "ColHeader",
        parent=styles["h3"],
        fontName=main_font,
        fontSize=10,
        leading=12,
        alignment=TA_CENTER,
    )
    normal_style = ParagraphStyle(
        "CellText",
        parent=styles["Normal"],
        fontName=main_font,
        fontSize=10,
        leading=12,  # Approx 1.2 * fontSize
        alignment=TA_LEFT,
    )
    min_edit_field_height = (
        2 * normal_style.leading + 0.2 * cm
    )  # Height for ~3 lines + padding

    # Margins
    margin_bottom = 1.5 * cm
    margin_left = 1.5 * cm
    margin_right = 1.5 * cm
    drawable_width = width - margin_left - margin_right

    # Add document title
    title_p = Paragraph(pdf_title_input, title_style)
    title_w, title_h = title_p.wrapOn(c, drawable_width, height)
    current_y_pos = height - 1.5 * cm
    title_p.drawOn(c, margin_left, current_y_pos - title_h)
    current_y_pos -= title_h + 0.5 * cm  # Space after title

    # Add wide edit field below title (first page only)
    form = c.acroForm  # Pega o form do canvas
    info_fields = [
        ("Detail1:", "field_tomador"),
        ("Detail2:", "field_credor"),
        ("Detail3:", "field_valor"),
        ("Detail4:", "field_processo"),
    ]

    label_font_size = 10
    field_height = 0.7 * cm
    field_width = drawable_width * 0.3  # Ocupa 30% da largura disponível
    label_width = drawable_width * 0.25

    # Espaço adicional entre o título e o primeiro campo
    current_y_pos -= 0.5 * cm

    for label_text, field_name in info_fields:
        # Rótulo
        c.setFont(main_font, label_font_size)
        c.drawString(margin_left, current_y_pos, label_text)

        # Campo editável
        form.textfield(
            name=field_name,
            tooltip=label_text.strip(":"),
            x=margin_left + label_width,
            y=current_y_pos - 0.1 * cm,
            width=field_width,
            height=field_height,
            value="",
            fontName="Helvetica",
            fontSize=10,
            borderColor=HexColor(0xAAAAAA),
            fillColor=HexColor(0xFFFFFF),
            textColor=HexColor(0x000000),
        )

        current_y_pos -= field_height + 0.3 * cm  # Espaço entre os campos

    margin_top_for_table = current_y_pos  # This is where the table headers will start

    rows = list(sheet.iter_rows(values_only=True))
    if not rows or len(rows) < 2:
        print("Erro: O arquivo Excel deve ter pelo menos 2 linhas (tipos e títulos).")
        return None

    col_types = [str(cell).lower() if cell is not None else "" for cell in rows[0]]
    col_titles = [str(cell) if cell is not None else "" for cell in rows[1]]
    data_rows = rows[2:]
    num_cols = len(col_types)

    if num_cols == 0:
        print("Erro: Nenhuma coluna encontrada na planilha.")
        return None

    fixed_checkbox_width = 1.0 * cm
    col_widths = [0] * num_cols
    fixed_width_allocated = 0
    num_dynamic_cols = 0

    for i, col_type in enumerate(col_types):
        if col_type == "checkbox":
            col_widths[i] = fixed_checkbox_width
            fixed_width_allocated += col_widths[i]
        else:
            num_dynamic_cols += 1

    remaining_width = drawable_width - fixed_width_allocated
    if num_dynamic_cols > 0 and remaining_width > num_dynamic_cols * (1 * cm):
        dynamic_col_width = remaining_width / num_dynamic_cols
        for i in range(num_cols):
            if col_types[i] != "checkbox":
                col_widths[i] = dynamic_col_width
    elif num_dynamic_cols > 0:
        print("Aviso: Espaço limitado para colunas dinâmicas. Distribuindo igualmente.")
        if remaining_width > 0:
            dynamic_col_width = remaining_width / num_dynamic_cols
            for i in range(num_cols):
                if col_types[i] != "checkbox":
                    col_widths[i] = dynamic_col_width
        else:
            print(
                "Aviso Crítico: Espaço insuficiente. Todas as colunas terão largura igual."
            )
            equal_width = drawable_width / num_cols
            col_widths = [equal_width] * num_cols
    elif num_dynamic_cols == 0 and fixed_width_allocated > drawable_width:
        print("Aviso: Colunas de checkbox excedem a largura. Ajustando.")
        scale = drawable_width / fixed_width_allocated
        for i in range(num_cols):
            col_widths[i] *= scale

    if any(w <= 0.1 * cm for w in col_widths):
        print("Aviso: Larguras de coluna inválidas. Usando larguras iguais.")
        col_widths = [drawable_width / num_cols] * num_cols

    min_row_height = 0.8 * cm
    checkbox_field_size = 0.5 * cm

    def draw_table_headers(start_y):
        c.setFont(main_font, 10)
        y = start_y
        max_header_h = 0
        header_paragraphs = []
        for i, title_text in enumerate(col_titles):
            temp_style = ParagraphStyle("TempHdr", parent=header_style)
            if col_types[i] == "checkbox":
                # Try to fit title, if not, make it very small or just use a symbol like "✓"
                # For now, allow wrap but it might look bad in a 1cm column.
                temp_style.alignment = TA_CENTER  # Center even checkbox title
                # title_text = "✓" # Alternative for very narrow checkbox columns
            p = Paragraph(title_text, temp_style)
            _w, p_h = p.wrapOn(c, col_widths[i] - 0.1 * cm, 3 * cm)
            header_paragraphs.append((p, p_h))
            max_header_h = max(max_header_h, p_h)

        actual_header_row_height = max(min_row_height, max_header_h + 0.2 * cm)

        for i, (p, p_h) in enumerate(header_paragraphs):
            x = margin_left + sum(col_widths[:i])
            p.drawOn(
                c, x + 0.05 * cm, y - (actual_header_row_height + p_h) / 2
            )  # Centered vertically
        y -= actual_header_row_height
        c.line(margin_left, y, width - margin_right, y)
        # form = c.acroForm # Not needed here, form is accessed via c.acroForm directly
        return y

    current_y = draw_table_headers(margin_top_for_table)

    for r_idx, data_row_values in enumerate(data_rows):
        max_cell_h_in_row = 0
        cell_contents = []

        for c_idx, cell_value in enumerate(data_row_values):
            if c_idx >= num_cols:
                continue
            col_type, cell_text = col_types[c_idx], str(
                cell_value if cell_value is not None else ""
            )
            cell_draw_width = col_widths[c_idx] - 0.2 * cm

            if col_type == "text":
                p = Paragraph(cell_text if cell_text else " ", normal_style)
                _w, p_h = p.wrapOn(c, cell_draw_width, drawable_width)
                cell_contents.append({"type": "text", "p": p, "h": p_h})
                max_cell_h_in_row = max(max_cell_h_in_row, p_h)
            elif col_type == "edit":
                cell_contents.append({"type": "edit", "h": min_edit_field_height})
                max_cell_h_in_row = max(max_cell_h_in_row, min_edit_field_height)
            elif col_type == "checkbox":
                cell_contents.append({"type": "checkbox", "h": checkbox_field_size})
                max_cell_h_in_row = max(max_cell_h_in_row, checkbox_field_size)
            else:
                p = Paragraph(f"? {cell_text}", normal_style)
                _w, p_h = p.wrapOn(c, cell_draw_width, drawable_width)
                cell_contents.append({"type": "unknown", "p": p, "h": p_h})
                max_cell_h_in_row = max(max_cell_h_in_row, p_h)

        dynamic_row_height = max(min_row_height, max_cell_h_in_row + 0.4 * cm)

        if current_y - dynamic_row_height < margin_bottom:
            c.showPage()
            # No main title or header_edit_field on subsequent pages
            current_y = draw_table_headers(height - 1.5 * cm)  # Start headers near top
            # c.acroForm needs to be used for the new page's form fields

        for c_idx, cell_info in enumerate(cell_contents):
            if c_idx >= num_cols:
                continue
            x = margin_left + sum(col_widths[:c_idx])
            content_h = cell_info["h"]
            # render_y is bottom-left y for paragraphs and fields, centered in row_height
            render_y = (
                current_y - dynamic_row_height + (dynamic_row_height - content_h) / 2
            )

            if cell_info["type"] == "text" or cell_info["type"] == "unknown":
                cell_info["p"].drawOn(c, x + 0.1 * cm, render_y)
            elif cell_info["type"] == "checkbox":
                c.acroForm.checkbox(
                    name=f"checkbox_{r_idx}_{c_idx}",
                    tooltip=f"{col_titles[c_idx]} {r_idx+1}",
                    x=x + (col_widths[c_idx] - checkbox_field_size) / 2,
                    y=current_y
                    - dynamic_row_height
                    + (dynamic_row_height - checkbox_field_size) / 2,
                    size=checkbox_field_size,
                    checked=False,
                    buttonStyle="check",
                    borderColor=HexColor(0x000000),
                    fillColor=HexColor(0xFFFFFF),
                    textColor=HexColor(0x000000),
                )
            elif cell_info["type"] == "edit":
                c.acroForm.textfield(
                    name=f"edit_{r_idx}_{c_idx}",
                    tooltip=f"{col_titles[c_idx]} {r_idx+1}",
                    x=x + 0.1 * cm,
                    y=current_y
                    - dynamic_row_height
                    + (dynamic_row_height - min_edit_field_height) / 2,
                    width=col_widths[c_idx] - 0.2 * cm,
                    height=min_edit_field_height,
                    value="",
                    fontName="Helvetica",
                    fontSize=7,
                    borderColor=HexColor(0xAAAAAA),
                    fillColor=HexColor(0xFFFFFF),
                    textColor=HexColor(0x000000),
                    fieldFlags="multiline",
                )

        current_y -= dynamic_row_height
        if r_idx < len(data_rows) - 1:
            c.line(margin_left, current_y, width - margin_right, current_y)

    c.save()
    print(f"PDF interativo '{pdf_file_name}' gerado com sucesso.")
    return pdf_file_name


if __name__ == "__main__":

    today_str = datetime.today().strftime("%d-%m-%Y")
    pdf_title_input_main = f"{base_title} {today_str}"

    create_interactive_pdf_v4(excel_file_input, pdf_title_input_main)
