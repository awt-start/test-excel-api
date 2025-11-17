# Excel科研经费通知生成系统

一个基于FastAPI的Excel文档自动生成系统，用于创建科研经费通知文件。

## 功能特性

- 📊 支持Excel模板渲染，可自定义横向/纵向模板
- 🔧 支持循环块渲染，可批量生成项目列表
- 📝 支持变量替换，动态填充通知内容
- 🛡️ 完善的数据验证和错误处理
- 🐳 支持Docker部署
- 📈 集成CI/CD工作流
- 📋 详细的日志记录，便于问题追踪

## 技术栈

- Python 3.14+
- FastAPI - Web框架
- Jinja2 - 模板引擎
- OpenPyXL - Excel文件处理
- Pydantic - 数据验证
- Uvicorn - ASGI服务器

## 快速开始

### 环境要求

- Python 3.14或更高版本
- pip或poetry

### 安装依赖

```bash
pip install -r requirements.txt
```

或使用poetry：

```bash
poetry install
```

### 启动开发服务器

```bash
uvicorn app.main:app --reload
```

服务器将在 http://localhost:8000 启动

### API文档

启动服务器后，可访问以下地址查看API文档：

- Swagger UI: http://localhost:8000/docs
- ReDoc: http://localhost:8000/redoc

## 使用Docker部署

### 构建镜像

```bash
docker build -t excel-fund-system .
```

### 运行容器

```bash
docker run -p 8000:8000 excel-fund-system
```

或使用docker-compose：

```bash
# 生产环境
docker-compose up -d production

# 开发环境
docker-compose up -d development
```

## API使用

### 渲染Excel通知

**POST /api/v1/excel/render**

请求体：

```json
{
  "template_type": "横向",
  "data": {
    "notice_no": "2025-KY-001",
    "date": "2025-11-18",
    "all_money": 150000.0,
    "signing_officer": "张三",
    "deputy1_dean": "李四",
    "top_leader": "王五",
    "finance_officer": "赵六",
    "deputy2_dean": "钱七",
    "research_handler": "孙八",
    "finance_handler": "周九",
    "projects": [
      {
        "project_code": "KY-2025-001",
        "project_name": "人工智能研究项目",
        "leader": "张教授",
        "department": "计算机学院",
        "source": "国家自然科学基金",
        "close_time": "2028-12-31",
        "money": 100000.0,
        "system_money": 5000.0,
        "public_consumption": 3000.0,
        "bank_name": "张教授",
        "open_bank": "中国工商银行",
        "bank_num": "6222021234567890123",
        "number": "110101199001011234"
      },
      {
        "project_code": "KY-2025-002",
        "project_name": "量子计算应用研究",
        "leader": "李教授",
        "department": "物理学院",
        "source": "科技部重点研发计划",
        "close_time": "2027-06-30",
        "money": 50000.0,
        "system_money": 2500.0,
        "public_consumption": 1500.0,
        "bank_name": "李教授",
        "open_bank": "中国建设银行",
        "bank_num": "6227001234567890456",
        "number": "110101198505056789"
      }
    ]
  }
}
```

### 健康检查

**GET /health**

返回系统健康状态信息

## 项目结构

```
excel-fund-two/
├── app/
│   ├── api/            # API路由
│   │   ├── endpoints/  # 具体接口实现
│   │   └── routes.py   # 路由配置
│   ├── services/       # 业务逻辑层
│   │   └── excel_renderer.py  # Excel渲染服务
│   ├── models/         # 数据模型
│   │   └── notice.py   # 通知数据模型
│   ├── templates/      # Excel模板目录
│   ├── main.py         # 应用入口
│   └── config.py       # 配置文件
├── tests/              # 测试文件
├── Dockerfile          # Docker配置
├── docker-compose.yml  # Docker Compose配置
├── pyproject.toml      # 项目依赖配置
├── requirements.txt    # 依赖清单
└── README.md           # 项目文档
```

## 开发指南

### 运行测试

```bash
pytest
```

### 代码检查

```bash
ruff check .
```

### 日志配置

系统使用Python标准日志模块，日志级别可在配置文件中设置。默认日志级别为INFO，包含以下日志模块：

- API请求日志
- Excel渲染过程日志
- 数据模型验证日志

## CI/CD

项目集成了GitHub Actions，包含以下工作流：

- **代码质量检查**：使用Ruff检查代码风格
- **单元测试**：运行pytest测试套件
- **Docker构建**：构建Docker镜像
- **部署**：自动部署到生产环境

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request！