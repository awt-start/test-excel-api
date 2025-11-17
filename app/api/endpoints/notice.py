# app/api/endpoints/notice.py
from fastapi import APIRouter, HTTPException, Response
from fastapi.responses import FileResponse

import os
import tempfile
import time
import logging
import uuid
from datetime import datetime
from app.models.notice import RenderRequest
from app.services.excel_renderer import render_excel_template

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
    # 生成请求唯一标识符
    request_id = str(uuid.uuid4())
    start_time = time.time()
    
    # 记录请求开始日志
    logger.info(f"[{request_id}] 开始处理Excel渲染请求")
    logger.info(f"[{request_id}] 模板类型: {request.template_type}")
    logger.info(f"[{request_id}] 通知编号: {request.data.notice_no if hasattr(request.data, 'notice_no') else '未提供'}")
    
    template_filename = TEMPLATE_MAP.get(request.template_type)
    if not template_filename:
        logger.error(f"[{request_id}] 模板类型 '{request.template_type}' 未找到")
        logger.info(f"[{request_id}] 可用模板类型: {list(TEMPLATE_MAP.keys())}")
        raise HTTPException(status_code=404, detail=f"模板类型 '{request.template_type}' 未找到。")

    logger.info(f"[{request_id}] 找到模板文件: {template_filename}")
    
    template_path = os.path.join(TEMPLATE_DIR, template_filename)
    if not os.path.exists(template_path):
        logger.error(f"[{request_id}] 模板文件不存在: {template_path}")
        raise HTTPException(status_code=404, detail=f"模板文件 '{template_path}' 不存在。")

    logger.info(f"[{request_id}] 模板文件验证通过")
    
    # Pydantic模型转字典，并处理别名
    try:
        context = request.data.model_dump()
        logger.info(f"[{request_id}] 数据模型转换成功，包含字段: {list(context.keys())}")
        
        # 记录项目数据摘要
        if 'projects' in context and context['projects']:
            logger.info(f"[{request_id}] 项目列表包含 {len(context['projects'])} 个项目")
            for i, project in enumerate(context['projects'][:3]):  # 只记录前3个项目
                logger.info(f"[{request_id}] 项目 {i+1}: {project.get('project_code', 'N/A')} - {project.get('project_name', 'N/A')}")
            if len(context['projects']) > 3:
                logger.info(f"[{request_id}] ... 还有 {len(context['projects']) - 3} 个项目")
        else:
            logger.warning(f"[{request_id}] 未找到项目数据")
            
    except Exception as e:
        logger.error(f"[{request_id}] 数据模型转换失败: {str(e)}")
        raise HTTPException(status_code=400, detail=f"数据格式错误: {str(e)}")

    tmp_file_path = None
    try:
        logger.info(f"[{request_id}] 开始渲染Excel模板")
        # 渲染Excel
        rendered_wb = render_excel_template(template_path, context)
        logger.info(f"[{request_id}] Excel模板渲染完成")

        # 将渲染后的工作簿保存到临时文件
        tmp_file_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp_file_path.close()  # 确保文件句柄关闭
        
        rendered_wb.save(tmp_file_path.name)
        logger.info(f"[{request_id}] 渲染文件已保存到临时路径: {tmp_file_path.name}")
        
        # 生成文件名
        filename = f"{request.template_type}_通知单_{context['notice_no']}.xlsx"
        logger.info(f"[{request_id}] 生成文件名: {filename}")
        
        # 计算处理时间
        process_time = time.time() - start_time
        logger.info(f"[{request_id}] 请求处理完成，耗时: {process_time:.3f}秒")
        
        # 返回文件给客户端
        response = FileResponse(
            path=tmp_file_path.name,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=filename
        )
        
        logger.info(f"[{request_id}] 开始发送响应文件")
        return response
        
    except Exception as e:
        logger.error(f"[{request_id}] Excel渲染过程中发生错误: {str(e)}", exc_info=True)
        
        # 清理临时文件（如果存在）
        if tmp_file_path and os.path.exists(tmp_file_path.name):
            try:
                os.remove(tmp_file_path.name)
                logger.info(f"[{request_id}] 已清理临时文件: {tmp_file_path.name}")
            except Exception as cleanup_error:
                logger.error(f"[{request_id}] 清理临时文件失败: {str(cleanup_error)}")
        
        # 计算失败时间
        process_time = time.time() - start_time
        logger.info(f"[{request_id}] 请求处理失败，耗时: {process_time:.3f}秒")
        
        raise HTTPException(status_code=500, detail=f"渲染Excel时发生错误: {str(e)}")
