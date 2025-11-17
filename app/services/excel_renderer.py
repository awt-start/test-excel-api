# app/services/excel_renderer.py
import os
import re
import openpyxl
from jinja2 import Environment, BaseLoader, TemplateSyntaxError
from openpyxl.utils import get_column_letter

# 正则表达式，用于匹配Jinja2的for循环块
FOR_LOOP_PATTERN = re.compile(r"{%\s*for\s+(\w+)\s+in\s+(\w+)\s*%}(.*){%\s*endfor\s*%}", re.DOTALL)

def render_excel_template(template_path: str, context: dict) -> openpyxl.Workbook:
    """
    渲染一个包含Jinja2语法的Excel模板。

    :param template_path: Excel模板文件路径。
    :param context: 用于渲染的上下文字典。
    :return: 渲染后的 openpyxl.Workbook 对象。
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    env = Environment(loader=BaseLoader())
    
    # --- 处理多行循环 ---
    # 我们需要先找到并处理循环行，因为它们会改变行数
    rows_to_process = []
    for row_idx, row in enumerate(ws.iter_rows(), 1):
        # 检查该行是否包含for循环
        # 我们将整行内容拼接成一个字符串来检查
        row_content = "".join([cell.value or "" for cell in row])
        if "{% for" in row_content and "{% endfor %}" in row_content:
            rows_to_process.append(row_idx)

    # 从后往前处理，避免插入行时影响后续的行索引
    for row_idx in reversed(rows_to_process):
        # 找到for循环的开始和结束标记所在的列
        start_col, end_col = None, None
        loop_var, list_name, loop_content = None, None, None

        for col_idx, cell in enumerate(ws[row_idx], 1):
            if cell.value and isinstance(cell.value, str):
                if "{% for" in cell.value:
                    start_col = col_idx
                if "{% endfor %}" in cell.value:
                    end_col = col_idx
        
        if not start_col or not end_col:
            continue # 跳过格式不正确的行

        # 提取循环内容
        loop_cell_value = ws.cell(row=row_idx, column=start_col).value
        match = FOR_LOOP_PATTERN.match(loop_cell_value)
        if not match:
            continue
            
        loop_var, list_name, loop_content = match.groups()
        project_list = context.get(list_name, [])

        # 删除原始模板行
        ws.delete_rows(row_idx)

        # 为列表中的每个项目创建一行
        for i, project_item in enumerate(project_list):
            # 插入新行
            ws.insert_rows(row_idx + i)
            
            # 创建一个临时的上下文，将循环变量（如'project'）指向当前项目
            temp_context = {loop_var: project_item}
            
            # 渲染这一行的每个单元格
            for col_idx in range(start_col, end_col + 1):
                original_cell_value = ws.cell(row=row_idx + i, column=col_idx).value
                if isinstance(original_cell_value, str) and "{{" in original_cell_value:
                    # 清理掉for/endfor标签，只保留内容
                    clean_value = original_cell_value.replace(f"{{% for {loop_var} in {list_name} %}}", "").replace("{% endfor %}", "").strip()
                    template = env.from_string(clean_value)
                    rendered_value = template.render(temp_context)
                    ws.cell(row=row_idx + i, column=col_idx, value=rendered_value)
                # 如果单元格内容只是for/endfor，则清空
                elif original_cell_value and ("{% for" in original_cell_value or "{% endfor %}" in original_cell_value):
                     ws.cell(row=row_idx + i, column=col_idx, value=None)


    # --- 处理普通变量 ---
    # 再次遍历所有单元格，处理简单的变量替换
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                # 跳过已经被循环处理过的复杂内容
                if "{%" in cell.value:
                    continue
                if "{{" in cell.value:
                    try:
                        template = env.from_string(cell.value)
                        rendered_value = template.render(context)
                        cell.value = rendered_value
                    except TemplateSyntaxError:
                        # 忽略语法错误，可能是未闭合的标签
                        pass
                        
    return wb