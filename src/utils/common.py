from typing import Dict, Any
from ..tools import pdf_tool, docx_tool, excel_tool, csv_tool


async def parse_page_range(page_range: str) -> set[int]:
    if not page_range or page_range.lower() == "all":
        return set()
    pages = set()
    for part in page_range.split(","):
        if "-" in part:
            start, end = map(int, part.split("-"))
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))
    return pages


async def call_tool(name: str, arguments: Dict[str, Any]) -> str:
    if name == "pdf_to_text":
        return await pdf_tool.pdf_to_text(arguments)
    elif name == "docx_to_markdown":
        return await docx_tool.docx_to_markdown(arguments)
    elif name == "excel_to_markdown":
        return await excel_tool.excel_to_markdown(arguments)
    elif name == "csv_to_markdown":
        return await csv_tool.csv_to_markdown(arguments)
    else:
        raise ValueError(f"未知工具: {name}")
