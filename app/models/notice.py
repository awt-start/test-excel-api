# app/models/notice.py
from pydantic import BaseModel, Field
from typing import List, Optional



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

class RenderRequest(BaseModel):
    template_type: str = Field(..., description="模板类型，如 '横向', '纵向'")
    data: NoticeData