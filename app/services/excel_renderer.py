# app/services/excel_renderer.py
import os
import re
import openpyxl
from jinja2 import Environment, BaseLoader, TemplateSyntaxError

# 正则表达式，用于匹配Jinja2的for循环的开始和结束
FOR_START_PATTERN = re.compile(r"{%\s*for\s+(\w+)\s+in\s+(\w+)\s*%}")
FOR_END_PATTERN = re.compile(r"{%\s*endfor\s*%}")

def _copy_style(source_cell, target_cell):
    """
    使用 openpyxl 内部机制，稳健地复制单元格样式。
    这是解决 'StyleArray' 问题的关键。
    """
    if source_cell.has_style:
        # 将源单元格的样式对象赋值给目标单元格
        # openpyxl 会自动处理样式的深拷贝
        target_cell._style = source_cell._style

def render_excel_template(template_path: str, context: dict) -> openpyxl.Workbook:
    """
    渲染一个包含Jinja2语法的Excel模板，支持多行循环并保留样式。
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    env = Environment(loader=BaseLoader())
    
    # --- 第一步：扫描并定位所有循环块 ---
    loop_blocks = []
    max_row = ws.max_row
    max_col = ws.max_column

    for row_idx in range(max_row, 0, -1):
        for col_idx in range(max_col, 0, -1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if not cell.value or not isinstance(cell.value, str):
                continue

            if FOR_END_PATTERN.search(cell.value):
                end_row, end_col = row_idx, col_idx
                start_row, start_col = None, None
                loop_var, list_name = None, None

                for r_idx in range(end_row, 0, -1):
                    for c_idx in range(max_col, 0, -1):
                        if r_idx == end_row and c_idx >= end_col:
                            continue
                        
                        start_cell = ws.cell(row=r_idx, column=c_idx)
                        if not start_cell.value or not isinstance(start_cell.value, str):
                            continue
                        
                        match = FOR_START_PATTERN.search(start_cell.value)
                        if match:
                            start_row, start_col = r_idx, c_idx
                            loop_var, list_name = match.groups()
                            break
                    if start_row:
                        break
                
                if start_row:
                    loop_blocks.append({
                        "start_row": start_row,
                        "end_row": end_row,
                        "start_col": start_col,
                        "end_col": end_col,
                        "loop_var": loop_var,
                        "list_name": list_name
                    })

    # --- 第二步：处理找到的循环块 ---
    for block in loop_blocks:
        start_row = block["start_row"]
        end_row = block["end_row"]
        start_col = block["start_col"]
        end_col = block["end_col"]
        loop_var = block["loop_var"]
        list_name = block["list_name"]

        project_list = context.get(list_name)
        if not project_list:
            ws.delete_rows(start_row, (end_row - start_row + 1))
            continue

        # 提取模板区域内的所有单元格值和样式
        template_rows_data = []
        for r_idx in range(start_row, end_row + 1):
            row_data = []
            for c_idx in range(start_col, end_col + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                # 存储单元格对象本身，以便后续复制样式
                row_data.append(cell)
            template_rows_data.append(row_data)
        
        ws.delete_rows(start_row, (end_row - start_row + 1))

        # --- 第三步：为每个项目创建新行并渲染 ---
        for i, project_item in enumerate(reversed(project_list)):
            ws.insert_rows(start_row, len(template_rows_data))
            
            temp_context = {loop_var: project_item}
            
            for r_offset, template_cells in enumerate(template_rows_data):
                for c_offset, template_cell in enumerate(template_cells):
                    current_cell = ws.cell(row=start_row + r_offset, column=start_col + c_offset)
                    original_value = template_cell.value

                    # *** 修正点：直接复制整个样式对象 ***
                    if template_cell.has_style:
                        _copy_style(template_cell, current_cell)

                    # 渲染值
                    if isinstance(original_value, str) and "{{" in original_value:
                        clean_value = FOR_START_PATTERN.sub("", original_value).strip()
                        clean_value = FOR_END_PATTERN.sub("", clean_value).strip()
                        
                        if clean_value:
                            try:
                                template = env.from_string(clean_value)
                                rendered_value = template.render(temp_context)
                                current_cell.value = rendered_value
                            except TemplateSyntaxError:
                                current_cell.value = original_value
                        else:
                            current_cell.value = None
                    else:
                        if isinstance(original_value, str) and ("{% for" in original_value or "{% endfor %}" in original_value):
                            current_cell.value = None
                        else:
                            current_cell.value = original_value

    # --- 第四步：处理页面中剩余的普通变量 ---
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "{{" in cell.value and "{%" not in cell.value:
                try:
                    template = env.from_string(cell.value)
                    rendered_value = template.render(context)
                    cell.value = rendered_value
                except TemplateSyntaxError:
                    pass
                        
    return wb