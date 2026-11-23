import sys
import subprocess
import tempfile
from pathlib import Path
from docx import Document

# 改用纯大写字母和数字，绝不会触发 Markdown 的下划线/星号等格式解析
TAB_PLACEHOLDER = "XTABCHAR999X"

def replace_in_paragraph(paragraph):
    """支持跨 Run 的精准替换，防止 Word 拆分文本碎片导致替换失败"""
    if TAB_PLACEHOLDER not in paragraph.text:
        return

    # 定位所有占位符在整段文字中的起始和结束索引
    while TAB_PLACEHOLDER in paragraph.text:
        full_text = paragraph.text
        start_idx = full_text.find(TAB_PLACEHOLDER)
        end_idx = start_idx + len(TAB_PLACEHOLDER)

        current_pos = 0
        for run in paragraph.runs:
            run_len = len(run.text)
            run_start = current_pos
            run_end = current_pos + run_len

            # 判断当前 run 是否与占位符有重叠
            if run_end > start_idx and run_start < end_idx:
                overlap_start = max(start_idx, run_start) - run_start
                overlap_end = min(end_idx, run_end) - run_start

                # 如果是包含占位符开头的第一个 run，将其重叠部分替换为 \t
                if run_start <= start_idx < run_end:
                    run.text = run.text[:overlap_start] + "\t" + run.text[overlap_end:]
                else:
                    # 后续被拆分的碎片直接清空重叠部分
                    run.text = run.text[:overlap_start] + run.text[overlap_end:]

            current_pos += run_len

def replace_placeholder_with_tab(docx_path: Path):
    """遍历 Word 文档的所有段落和表格"""
    doc = Document(str(docx_path))

    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph)

    doc.save(str(docx_path))

def convert_md_to_docx(input_arg: str):
    input_path = Path(input_arg).resolve()
    
    if not input_path.suffix:
        input_path = input_path.with_suffix(".md")

    if not input_path.exists():
        print(f"错误：找不到文件 -> {input_path}")
        sys.exit(1)

    output_path = input_path.with_suffix(".docx")

    # 替换全角空格 (\u3000) 为纯字母占位符
    content = input_path.read_text(encoding="utf-8")
    processed_content = content.replace("\u3000", TAB_PLACEHOLDER)

    with tempfile.NamedTemporaryFile(mode="w", suffix=".md", encoding="utf-8", delete=False) as tmp_md:
        tmp_md.write(processed_content)
        tmp_md_path = Path(tmp_md.name)

    try:
        cmd = [
            "pandoc",
            str(tmp_md_path),
            "-f", "markdown+hard_line_breaks+tex_math_dollars+tex_math_single_backslash+pipe_tables+grid_tables",
            "--highlight-style=tango",
            "--toc",
            "--toc-depth=3",
            "-o", str(output_path)
        ]
        
        print(f"正在转换: {input_path.name} -> {output_path.name} ...")
        subprocess.run(cmd, check=True)

        replace_placeholder_with_tab(output_path)
        print(f"转换成功！输出文件：{output_path}")

    except subprocess.CalledProcessError as e:
        print(f"Pandoc 执行失败，错误码: {e.returncode}")
    except FileNotFoundError:
        print("错误：未找到 pandoc 命令，请确认已安装 Pandoc 并添加到环境变量 PATH 中。")
    finally:
        if tmp_md_path.exists():
            tmp_md_path.unlink()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python pandoc01.py <markdown文件路径>")
        sys.exit(1)

    convert_md_to_docx(sys.argv[1])




