import re
import os
from copy import copy
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

VAR_PATTERN = re.compile(r"{{\s*([\w\.]+)\s*}}")


def resolve_var(name, context):
    """支持 a.b.c 形式的取值"""
    parts = name.split(".")
    val = context
    for p in parts:
        val = val[p]
    return val


def render_excel(template_path, output_path, data):
    wb = load_workbook(template_path)

    # ---- 1. 处理普通文本占位符 {{var}} ----
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    matches = VAR_PATTERN.findall(cell.value)
                    if not matches:
                        continue
                    new_val = cell.value
                    for varname in matches:
                        new_val = new_val.replace(
                            "{{"+varname+"}}",
                            str(resolve_var(varname, data))
                        )
                    cell.value = new_val

    # ---- 2. 处理表格循环：包含 item.xxx 的行视为模板 ----
    for ws in wb.worksheets:
        template_rows = []
        for row in ws.iter_rows():
            row_str = "".join([str(c.value) for c in row if c.value])
            if "item." in row_str:    # 简单识别方式
                template_rows.append(row[0].row)

        for row_idx in reversed(template_rows):
            template_row = ws[row_idx]
            ws.delete_rows(row_idx)

            items = data["items"]
            for item in reversed(items):
                ws.insert_rows(row_idx)
                for c in template_row:
                    new_cell = ws.cell(row=row_idx, column=c.col_idx)
                    new_cell.value = c.value
                    new_cell.font = copy(c.font)
                    new_cell.border = copy(c.border)
                    new_cell.alignment = copy(c.alignment)
                    new_cell.fill = copy(c.fill)
                    new_cell.number_format = copy(c.number_format)

                    # 替换 {{item.xxx}}
                    if isinstance(new_cell.value, str):
                        matches = VAR_PATTERN.findall(new_cell.value)
                        for varname in matches:
                            if varname.startswith("item."):
                                field = varname.split(".", 1)[1]
                                new_cell.value = new_cell.value.replace(
                                    "{{"+varname+"}}", str(item[field])
                                )

    # ---- 3. 处理图片插入 images.logo -> A1 等 ----
    if "images" in data:
        for img_name, path in data["images"].items():
            if not isinstance(path, str):
                continue
            if not os.path.exists(path):
                continue
            # 简单约定：在某个单元格写 {{ logo }} 的位置插入图片
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str) and f"{{{{{img_name}}}}}" in cell.value:
                            ws.add_image(XLImage(path), cell.coordinate)
                            cell.value = None  # 清除占位符

    wb.save(output_path)
