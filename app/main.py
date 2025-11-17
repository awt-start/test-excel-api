# app/main.py
from fastapi import FastAPI
from app.api.endpoints import notice
import logging

app = FastAPI(
    title="Excel渲染API",
    description="一个用于根据模板和数据渲染Excel文件的服务。",
    version="1.0.0",
)
# 定义logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

app.include_router(notice.router, prefix="/api/v1/notices", tags=["通知单"])


@app.get("/")
def read_root():
    return {"message": "欢迎使用Excel渲染API，请访问 /docs 查看文档。"}
