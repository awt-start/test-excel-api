# app/models/notice.py
import logging
from datetime import datetime
from pydantic import BaseModel, Field, model_validator
from typing import List, Optional

# 配置日志
logger = logging.getLogger(__name__)



class ProjectInfo(BaseModel):
    project_code: Optional[str] = Field(None, description="项目编码")
    project_name: Optional[str] = Field(None, description="项目名称")
    leader: Optional[str] = Field(None, description="负责人")
    department: Optional[str] = Field(None, description="所在部门/单位")
    source: Optional[str] = Field(None, description="经费/项目来源")
    close_time: Optional[str] = Field(None, description="经费截止时间")
    money: Optional[float] = Field(None, description="拨款数额/经费")
    system_money: Optional[float] = Field(None, description="管理费")
    public_consumption: Optional[float] = Field(None, description="公共物耗")
    bank_name: Optional[str] = Field(None, description="银行用户名")
    open_bank: Optional[str] = Field(None, description="开户行")
    bank_num: Optional[str] = Field(None, description="银行帐号")
    number: Optional[str] = Field(None, description="身份证号")
    
    @model_validator(mode='after')
    def log_project_info(self):
        """记录项目信息验证完成后的日志"""
        project_identifier = self.project_code or self.project_name or "未知项目"
        logger.info(f"[数据模型] 项目信息验证完成: {project_identifier}")
        logger.debug(f"[数据模型] 项目详情: {self.model_dump(exclude_none=True)}")
        return self

class NoticeData(BaseModel):
    notice_no: str = Field(..., description="通知编号")
    date: str = Field(..., description="通知日期")
    all_money: Optional[float] = Field(None, description="项目总经费")
    signing_officer: Optional[str] = Field(None, description="科研部长签字")
    deputy1_dean: Optional[str] = Field(None, description="分管院长审核(科研)")
    top_leader: Optional[str] = Field(None, description="院长审批")
    finance_officer: Optional[str] = Field(None, description="财务部长签字")
    deputy2_dean: Optional[str] = Field(None, description="分管院长审核(财务)")
    research_handler: Optional[str] = Field(None, description="科研部经手人")
    finance_handler: Optional[str] = Field(None, description="财务部经手人")
    
    projects: List[ProjectInfo] = Field(default_factory=list, description="项目列表")
    
    @model_validator(mode='after')
    def log_notice_data(self):
        """记录通知数据验证完成后的日志"""
        logger.info(f"[数据模型] 通知数据验证完成: {self.notice_no} (日期: {self.date})")
        logger.info(f"[数据模型] 包含{len(self.projects)}个项目")
        
        # 验证项目总经费是否与各项目经费总和匹配
        if self.all_money is not None and self.projects:
            total_project_money = sum(project.money or 0 for project in self.projects)
            if abs(self.all_money - total_project_money) > 0.01:  # 允许0.01的浮点数误差
                logger.warning(f"[数据模型] 总经费与项目经费总和不匹配: {self.all_money} != {total_project_money}")
            else:
                logger.debug(f"[数据模型] 总经费与项目经费总和匹配: {self.all_money}")
                
        logger.debug(f"[数据模型] 通知详情: {self.model_dump(exclude_none=True, exclude={'projects'})}")
        return self

class RenderRequest(BaseModel):
    template_type: str = Field(..., description="模板类型，如 '横向', '纵向'")
    data: NoticeData
    
    @model_validator(mode='after')
    def log_render_request(self):
        """记录渲染请求验证完成后的日志"""
        logger.info(f"[数据模型] 渲染请求验证完成: 模板类型='{self.template_type}', 通知编号='{self.data.notice_no}'")
        logger.debug(f"[数据模型] 渲染请求详情: {self.model_dump(exclude_none=True)}")
        return self