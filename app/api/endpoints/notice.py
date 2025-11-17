# app/api/endpoints/notice.py
from fastapi import APIRouter, HTTPException, Response
from fastapi.responses import FileResponse

import os
import tempfile
from app.models.notice import RenderRequest
from app.services.excel_renderer import render_excel_template

router = APIRouter()

# 模板类型到文件名的映射
TEMPLATE_MAP = {
    "横向": "template_横向.xlsx",
    "纵向": "template_纵向.xlsx",
    "协同创新专项": "template_协同创新专项.xlsx",
    "协同创新专项精品": "template_协同创新专项精品.xlsx",
}
TEMPLATE_DIR = "app/templates"

@router.post("/render")
async def render_notice(request: RenderRequest):
    """
    接收渲染请求，生成Excel文件并返回。
    """
    template_filename = TEMPLATE_MAP.get(request.template_type)
    if not template_filename:
        raise HTTPException(status_code=404, detail=f"模板类型 '{request.template_type}' 未找到。")

    template_path = os.path.join(TEMPLATE_DIR, template_filename)
    if not os.path.exists(template_path):
        raise HTTPException(status_code=404, detail=f"模板文件 '{template_path}' 不存在。")

    # Pydantic模型转字典，并处理别名
    context =  request.data.model_dump() 

    try:
        # 渲染Excel
        rendered_wb = render_excel_template(template_path, context)

        # 将渲染后的工作簿保存到临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            rendered_wb.save(tmp.name)
            
            # 返回文件给客户端
            return FileResponse(
                path=tmp.name,
                media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                filename=f"{request.template_type}_通知单_{context['notice_no']}.xlsx"
            )
    except Exception as e:
        # 清理临时文件（如果存在）
        if 'tmp' in locals() and os.path.exists(tmp.name):
            os.remove(tmp.name)
        raise HTTPException(status_code=500, detail=f"渲染Excel时发生错误: {str(e)}")
