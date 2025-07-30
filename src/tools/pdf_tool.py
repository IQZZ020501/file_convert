import pdfplumber
from pathlib import Path
import logging
from ..utils.common import parse_page_range

logger = logging.getLogger("file_to_markdown")


async def pdf_to_text(arguments: dict) -> str:
    try:
        file_path = arguments.get("file_path")
        page_range = arguments.get("page_range", "all")
        if not file_path:
            return "错误: 需要提供PDF文件路径"

        pdf_path = Path(file_path)
        if not pdf_path.exists():
            return f"错误: 文件不存在: {file_path}"

        if pdf_path.suffix.lower() != ".pdf":
            return "错误: 文件必须是PDF格式"

        pages_to_extract = await parse_page_range(page_range)

        extracted_text = []
        page_count = 0

        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            for i, page in enumerate(pdf.pages, 1):
                if pages_to_extract and i not in pages_to_extract:
                    continue
                text = page.extract_text()
                if text:
                    extracted_text.append(f"=== 第 {i} 页 ===\n{text}\n")
                    page_count += 1

        return f"成功提取文本！\n\n文件: {file_path}\n总页数: {total_pages}\n提取页数: {page_count}\n\n{''.join(extracted_text)}"
    except Exception as e:
        logger.error(f"PDF转文字失败: {e}")
        return f"错误: {e}"
