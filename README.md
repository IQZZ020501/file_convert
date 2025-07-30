# 文档转换工具 (File to Markdown Converter)

一个基于 MCP (Model Context Protocol) 的文档转换工具，支持将多种格式的文档转换为 Markdown 格式。

## 功能特性

✨ **支持的文档格式**
- 📄 PDF 文件 (.pdf)
- 📝 Word 文档 (.docx, .doc)
- 📊 Excel 表格 (.xlsx, .xls)
- 📋 CSV 文件 (.csv)

✨ **主要功能**
- PDF 文本提取并转换为 Markdown
- Word 文档格式保留转换（标题、加粗、斜体、表格）
- Excel 多工作表转换为 Markdown 表格
- CSV 数据转换为 Markdown 表格
- 支持页面范围指定（PDF）
- 支持工作表选择（Excel）

## 系统要求

- Python >= 3.11
- Windows/macOS/Linux

## 安装方法

### 1. 克隆项目
```bash
git clone <repository-url>
cd file_convert
```

### 2. 安装依赖
```bash
pip install -e .
```

或者使用 uv：
```bash
uv sync
```

## 使用方法

### 作为 MCP 服务器

1. 启动 MCP 服务器：
```bash
uv run -m src.server
```

2. 通过 MCP 客户端调用工具

### 支持的工具

```json
{
  "mcpServers": {
    "file_convert": {
      "command": "uv",
      "args": [
        "--directory",
        "D:\\file_convert\\src",
        "run",
        "server.py"
      ]
    }
  }
}

```

## 转换示例

### Word 文档转换效果
**输入：** Word 文档包含标题、段落、表格
**输出：** 
```markdown
# 主标题
## 副标题
这是一个**加粗文本**和*斜体文本*的段落。

| 列1 | 列2 | 列3 |
| --- | --- | --- |
| 数据1 | 数据2 | 数据3 |
```

### Excel 转换效果
**输入：** Excel 表格数据
**输出：**
```markdown
# Sheet1

| 姓名 | 年龄 | 城市 |
| --- | --- | --- |
| 张三 | 25 | 北京 |
| 李四 | 30 | 上海 |
```

## 项目结构

```
pdf_convert/
├── pyproject.toml          # 项目配置文件
├── README.md              # 项目说明文档
├── uv.lock               # 依赖锁定文件
└── src/
    └── mcp_server_pdf_convert/
        └── server.py     # MCP 服务器主程序
```

## 依赖库

- **mcp[cli]** - Model Context Protocol 框架
- **pdfplumber** - PDF 文本提取
- **python-docx** - Word 文档处理
- **openpyxl** - Excel 文件处理
- **xlrd** - 旧版 Excel 文件支持
- **pandas** - 数据处理

## 错误处理

工具会自动处理以下错误情况：
- 文件不存在
- 格式不支持
- 文件损坏
- 权限不足
- 空文档处理

所有错误都会返回详细的错误信息，便于调试。

## 开发和贡献

### 本地开发
```bash
# 安装开发依赖
uv sync --dev

# 运行测试
python -m pytest

# 代码格式化
black src/
isort src/
```

### 添加新功能
1. 在 `server.py` 中添加新的工具函数
2. 在 `call_tool` 函数中注册新工具
3. 添加相应的测试用例
4. 更新文档

## 许可证

本项目采用 MIT 许可证，详见 LICENSE 文件。

## 更新日志

### v1.0.0
- ✅ 支持 PDF 转文本
- ✅ 支持 Word 文档转 Markdown
- ✅ 支持 Excel 转 Markdown 表格
- ✅ 支持 CSV 转 Markdown 表格
- ✅ 支持页面范围和工作表选择
- ✅ 完整的错误处理机制

## 联系方式

如有问题或建议，请提交 Issue 或 Pull Request。
