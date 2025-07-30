import logging
from pathlib import Path
import pandas as pd
from typing import Dict, Any

logger = logging.getLogger("file_to_markdown")


async def csv_to_markdown(arguments: Dict[str, Any]) -> str:
    """将CSV文件转换为Markdown格式"""
    try:
        file_path = arguments.get("file_path")
        encoding = arguments.get("encoding", "utf-8")
        delimiter = arguments.get("delimiter", ",")

        if not file_path:
            return "错误: 需要提供CSV文件路径"

        csv_path = Path(file_path)
        if not csv_path.exists():
            return f"错误: 文件不存在: {file_path}"

        if not csv_path.suffix.lower() == '.csv':
            return "错误: 文件必须是CSV格式"

        encodings_to_try = [encoding, 'utf-8', 'gbk', 'gb2312', 'latin-1']
        df = None
        used_encoding = None

        for enc in encodings_to_try:
            try:
                df = pd.read_csv(csv_path, encoding=enc, delimiter=delimiter)
                used_encoding = enc
                break
            except UnicodeDecodeError:
                continue
            except Exception as e:
                if enc == encodings_to_try[-1]:
                    raise e
                continue

        if df is None:
            return "错误: 无法读取CSV文件，请检查文件编码"

        df = df.fillna("")

        markdown_content = [f"# CSV文件转换结果", f"**文件**: {file_path}", f"**编码**: {used_encoding}",
                            f"**行数**: {len(df)}", f"**列数**: {len(df.columns)}", ""]

        if len(df) == 0:
            markdown_content.append("*CSV文件为空*")
        else:
            headers = [str(col) for col in df.columns]
            markdown_content.append("| " + " | ".join(headers) + " |")

            separator = "| " + " | ".join(["---"] * len(headers)) + " |"
            markdown_content.append(separator)

            for _, row in df.iterrows():
                row_data = [str(value).replace("|", "\\|").replace("\n", " ") for value in row]
                markdown_content.append("| " + " | ".join(row_data) + " |")

        result = "\n".join(markdown_content)
        return f"成功转换为Markdown！\n\n{result}"

    except Exception as e:
        logger.error(f"CSV转Markdown过程中发生错误: {str(e)}")
        return f"错误: {str(e)}"
