import logging

from mcp.server.fastmcp import FastMCP

from src.utils.common import call_tool

mcp = FastMCP(
    name="file_to_markdown",
    description="将PDF、Word、Excel、csv 转换为Markdown格式的工具",
    version="1.0.0",
)

logging.basicConfig(level=logging.INFO)


@mcp.tool(name="pdf_to_text", description="将PDF文件转换为文本")
async def pdf_convert(file_path: str, page_range: str = "all") -> str:
    return await call_tool("pdf_to_text", {"file_path": file_path, "page_range": page_range})


@mcp.tool(name="docx_to_markdown", description="将DOC/DOCX文件转换为Markdown格式")
async def docx_convert(file_path: str) -> str:
    return await call_tool("docx_to_markdown", {"file_path": file_path})


@mcp.tool(name="excel_to_markdown", description="将Excel文件转换为Markdown格式")
async def excel_convert(file_path: str, sheet_name: str = None) -> str:
    return await call_tool("excel_to_markdown", {"file_path": file_path, "sheet_name": sheet_name})


@mcp.tool(name="csv_to_markdown", description="将CSV文件转换为Markdown格式")
async def csv_convert(file_path: str, encoding: str = "utf-8", delimiter: str = ",") -> str:
    return await call_tool("csv_to_markdown", {
        "file_path": file_path,
        "encoding": encoding,
        "delimiter": delimiter
    })


if __name__ == "__main__":
    mcp.run(transport="stdio")
