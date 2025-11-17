# app/models/notice.py
from pydantic import BaseModel, Field
from typing import List, Optional

# 定义一个通用的项目信息模型
# 使用 Optional 允许某些字段为空，以适应不同模板
class ProjectInfo(BaseModel):
    project_code: Optional[str] = Field(None, alias="项目编码")
    project_name: Optional[str] = Field(None, alias="项目名称")
    leader: Optional[str] = Field(None, alias="负责人")
    department: Optional[str] = Field(None, alias="所在部门/单位")
    source: Optional[str] = Field(None, alias="经费/项目来源")
    close_time: Optional[str] = Field(None, alias="经费截止时间")
    money: Optional[float] = Field(None, alias="拨款数额/经费")
    system_money: Optional[float] = Field(None, alias="管理费")
    public_consumption: Optional[float] = Field(None, alias="公共物耗")
    bank_name: Optional[str] = Field(None, alias="银行用户名")
    open_bank: Optional[str] = Field(None, alias="开户行")
    bank_num: Optional[str] = Field(None, alias="银行帐号")
    number: Optional[str] = Field(None, alias="身份证号")

# 定义顶层通知信息模型
class NoticeData(BaseModel):
    notice_no: str = Field(..., alias="编号")
    date: str = Field(..., alias="日期")
    all_money: Optional[float] = Field(None, alias="项目总经费")
    signing_officer: Optional[str] = Field(None, alias="科研部长签字")
    deputy1_dean: Optional[str] = Field(None, alias="分管院长审核(科研)")
    top_leader: Optional[str] = Field(None, alias="院长审批")
    finance_officer: Optional[str] = Field(None, alias="财务部长签字")
    deputy2_dean: Optional[str] = Field(None, alias="分管院长审核(财务)")
    research_handler: Optional[str] = Field(None, alias="科研部经手人")
    finance_handler: Optional[str] = Field(None, alias="财务部经手人")
    
    # 使用列表，因为一个通知单可能包含多个项目
    projects: List[ProjectInfo] = Field(default_factory=list, alias="项目列表")

# 定义API请求的最终模型
class RenderRequest(BaseModel):
    template_type: str = Field(..., description="模板类型，如 '横向', '纵向', '协同创新专项'")
    data: NoticeData