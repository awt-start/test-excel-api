#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel渲染API测试脚本
"""

import json
import requests

# API基础URL
BASE_URL = "http://localhost:8000"

def test_root():
    """测试根路径"""
    print("=== 测试根路径 ===")
    response = requests.get(f"{BASE_URL}/")
    print(f"状态码: {response.status_code}")
    print(f"响应: {response.json()}")
    print()

def test_render_api():
    """测试Excel渲染API"""
    print("=== 测试Excel渲染接口 ===")
    
    # 构建测试数据
    test_data = {
        "template_type": "横向",
        "data": {
            "notice_no": "TEST001",
            "date": "2024-11-17",
            "all_money": 100000.0,
            "signing_officer": "张部长",
            "deputy1_dean": "李院长",
            "top_leader": "王院长",
            "finance_officer": "赵部长",
            "deputy2_dean": "钱院长",
            "research_handler": "小张",
            "finance_handler": "小李",
            "projects": [
                {
                    "project_code": "PROJ001",
                    "project_name": "测试项目A",
                    "leader": "张三",
                    "department": "计算机学院",
                    "source": "横向合作",
                    "close_time": "2024-12-31",
                    "money": 50000.0,
                    "system_money": 5000.0,
                    "public_consumption": 2000.0
                },
                {
                    "project_code": "PROJ002", 
                    "project_name": "测试项目B",
                    "leader": "李四",
                    "department": "信息学院",
                    "source": "企业合作",
                    "close_time": "2025-06-30",
                    "money": 50000.0,
                    "system_money": 5000.0,
                    "public_consumption": 3000.0
                }
            ]
        }
    }
    
    try:
        response = requests.post(
            f"{BASE_URL}/api/v1/notices/render",
            json=test_data,
            headers={"Content-Type": "application/json"}
        )
        
        print(f"状态码: {response.status_code}")
        
        if response.status_code == 200:
            # 保存下载的Excel文件
            filename = f"test_rendered_{test_data['data']['notice_no']}.xlsx"
            with open(filename, "wb") as f:
                f.write(response.content)
            print(f"✅ Excel文件已成功生成并保存为: {filename}")
            print(f"文件大小: {len(response.content)} 字节")
        else:
            print(f"❌ 渲染失败")
            print(f"错误信息: {response.text}")
            
    except Exception as e:
        print(f"❌ 请求异常: {str(e)}")
    
    print()

def test_template_types():
    """测试不同的模板类型"""
    print("=== 测试不同模板类型 ===")
    
    template_types = ["横向", "纵向", "协同创新专项", "协同创新专项精品"]
    
    for template_type in template_types:
        print(f"测试模板类型: {template_type}")
        
        test_data = {
            "template_type": template_type,
            "data": {
                "notice_no": f"TEMP{template_type}001",
                "date": "2024-11-17",
                "projects": [
                    {
                        "project_code": f"PROJ{template_type}001",
                        "project_name": f"{template_type}测试项目",
                        "leader": "测试负责人",
                        "department": "测试部门",
                        "source": "测试来源",
                        "close_time": "2024-12-31",
                        "money": 10000.0,
                        "system_money": 1000.0,
                        "public_consumption": 500.0
                    }
                ]
            }
        }
        
        try:
            response = requests.post(
                f"{BASE_URL}/api/v1/notices/render",
                json=test_data,
                headers={"Content-Type": "application/json"}
            )
            
            if response.status_code == 200:
                filename = f"test_{template_type}_template.xlsx"
                with open(filename, "wb") as f:
                    f.write(response.content)
                print(f"✅ {template_type}模板渲染成功，保存为: {filename}")
            else:
                print(f"❌ {template_type}模板渲染失败: {response.status_code} - {response.text}")
                
        except Exception as e:
            print(f"❌ {template_type}模板测试异常: {str(e)}")
        
        print()

if __name__ == "__main__":
    print("Excel渲染API接口测试开始...\n")
    
    test_root()
    test_render_api()
    test_template_types()
    
    print("测试完成！")