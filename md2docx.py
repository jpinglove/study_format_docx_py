import sys
import argparse
import subprocess
import tempfile
from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Word 中文标准字号：五号 = 10.5 磅（OpenXML 中以半磅为单位，即 21）
WUHAO_PT = 10.5
WUHAO_HALF_PT = 21

# 改用纯大写字母和数字，绝不会触发 Markdown 的下划线/星号等格式解析
TAB_PLACEHOLDER = "XTABCHAR999X"

# '''
# # 1. 直接将 Markdown 转成最终排版好的同名 docx（c10.md -> c10.docx）
# python md2docx.py ./c10.md
# 
# # 2. 不写后缀也行，自动按 .md 查找
# python md2docx.py C:\AIScripts\exam\exame_papers\c10
# 
# # 3. 指定输出文件名
# python md2docx.py ./c10.md -o ./c10_print.docx
# 
# # 4. 如果不需要生成目录，加上 --no-toc
# python md2docx.py ./c10.md --no-toc
# 
# # 5. 如果输入本身就是 .docx，也可直接复用排版逻辑
# python md2docx.py ./c10.docx
# '''


# Office Math (OMML) 命名空间
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


# 常见纯文本项目符号前缀（兼容 Markdown/HTML 转 docx 未生成 w:numPr 的情况）
LITERAL_BULLET_PREFIXES = (
    "•", "◦", "▪", "▫", "○", "●", "◆", "◇", "■", "□", "➢", "➤", "✓", "✔"
)


def m_qn(tag: str) -> str:
    return f"{{{MATH_NS}}}{tag}"


# ==========================================
# 第一部分：Markdown 占位符与 Pandoc 转换
# ==========================================

def replace_in_paragraph(paragraph):
    """支持跨 Run 的精准替换，防止 Word 拆分文本碎片导致替换失败"""
    if TAB_PLACEHOLDER not in paragraph.text:
        return

    while TAB_PLACEHOLDER in paragraph.text:
        full_text = paragraph.text
        start_idx = full_text.find(TAB_PLACEHOLDER)
        end_idx = start_idx + len(TAB_PLACEHOLDER)

        current_pos = 0
        for run in paragraph.runs:
            run_len = len(run.text)
            run_start = current_pos
            run_end = current_pos + run_len

            if run_end > start_idx and run_start < end_idx:
                overlap_start = max(start_idx, run_start) - run_start
                overlap_end = min(end_idx, run_end) - run_start

                if run_start <= start_idx < run_end:
                    run.text = run.text[:overlap_start] + "\t" + run.text[overlap_end:]
                else:
                    run.text = run.text[:overlap_start] + run.text[overlap_end:]

            current_pos += run_len


def replace_placeholder_with_tab(doc: Document):
    """遍历 Word 文档 XML 树中的所有段落（含表格、嵌套表格、内容控件），还原制表符"""
    for p_el in doc.element.body.iter(qn("w:p")):
        paragraph = Paragraph(p_el, doc._body)
        replace_in_paragraph(paragraph)


def run_pandoc_to_docx(md_path: Path, raw_docx_path: Path, with_toc: bool = True):
    """将全角空格替换为占位符后，调用 pandoc 转为临时 docx"""
    content = md_path.read_text(encoding="utf-8")
    processed_content = content.replace("\u3000", TAB_PLACEHOLDER)

    tmp_md_path = raw_docx_path.with_name("source_processed.md")
    tmp_md_path.write_text(processed_content, encoding="utf-8")

    cmd = [
        "pandoc",
        str(tmp_md_path),
        "-f", "markdown+hard_line_breaks+tex_math_dollars+tex_math_single_backslash+pipe_tables+grid_tables",
        "--highlight-style=tango",
    ]
    if with_toc:
        cmd.extend(["--toc", "--toc-depth=3"])
    cmd.extend(["-o", str(raw_docx_path)])

    print(f"正在通过 Pandoc 转换: {md_path.name} -> {raw_docx_path.name} ...")
    subprocess.run(cmd, check=True)


# ==========================================
# 第二部分：DOCX 排版、字号、对齐与页码处理
# ==========================================

def get_doc_default_font_size_pt(doc):
    """读取文档 docDefaults 中的默认字号（单位：磅）。"""
    styles_el = doc.styles.element
    doc_defaults = styles_el.find(qn("w:docDefaults"))
    if doc_defaults is None:
        return None
    rpr_default = doc_defaults.find(qn("w:rPrDefault"))
    if rpr_default is None:
        return None
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        return None
    sz = rpr.find(qn("w:sz"))
    if sz is not None and sz.get(qn("w:val")):
        return int(sz.get(qn("w:val"))) / 2.0
    return None


def fix_doc_defaults_font_size(doc):
    """如果文档默认字号大于五号(10.5磅)，将其改为五号。"""
    styles_el = doc.styles.element
    doc_defaults = styles_el.find(qn("w:docDefaults"))
    if doc_defaults is None:
        return False
    rpr_default = doc_defaults.find(qn("w:rPrDefault"))
    if rpr_default is None:
        return False
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        return False

    changed = False
    for tag in ("w:sz", "w:szCs"):
        sz = rpr.find(qn(tag))
        if sz is not None and sz.get(qn("w:val")):
            if int(sz.get(qn("w:val"))) > WUHAO_HALF_PT:
                sz.set(qn("w:val"), str(WUHAO_HALF_PT))
                changed = True
    return changed


def set_math_default_left(doc) -> bool:
    """将 word/settings.xml 中的全局公式默认对齐 (<m:defJc>) 设为 left。"""
    settings_el = doc.settings.element
    math_pr = settings_el.find(m_qn("mathPr"))
    if math_pr is None:
        return False

    def_jc = math_pr.find(m_qn("defJc"))
    if def_jc is not None:
        def_jc.attrib.clear()
        def_jc.set(m_qn("val"), "left")
        return True

    def_jc = math_pr.makeelement(m_qn("defJc"))
    def_jc.set(m_qn("val"), "left")
    after_tags = {
        m_qn("preSp"),
        m_qn("postSp"),
        m_qn("interSp"),
        m_qn("intraSp"),
        m_qn("wrapIndent"),
        m_qn("wrapRight"),
        m_qn("intLim"),
        m_qn("naryLim"),
    }
    insert_idx = len(math_pr)
    for idx, child in enumerate(math_pr):
        if child.tag in after_tags:
            insert_idx = idx
            break
    math_pr.insert(insert_idx, def_jc)
    return True


def get_style_font_size_pt(style):
    """沿着样式继承链向上查找字号（单位：磅）。"""
    visited = set()
    curr = style
    while curr is not None and curr.name not in visited:
        visited.add(curr.name)
        if curr.font is not None and curr.font.size is not None:
            return curr.font.size.pt
        curr = getattr(curr, "base_style", None)
    return None


def get_effective_font_size_pt(run, paragraph, doc):
    """获取 Run 实际生效的字号。"""
    if run.font is not None and run.font.size is not None:
        return run.font.size.pt

    if run.style is not None:
        size = get_style_font_size_pt(run.style)
        if size is not None:
            return size

    if paragraph.style is not None:
        size = get_style_font_size_pt(paragraph.style)
        if size is not None:
            return size

    if "Normal" in doc.styles:
        size = get_style_font_size_pt(doc.styles["Normal"])
        if size is not None:
            return size

    return get_doc_default_font_size_pt(doc)


def has_active_numpr(pPr_element):
    """判断 pPr 元素中是否包含有效的编号/项目符号配置 (numId != 0)。"""
    if pPr_element is None:
        return False
    numPr = pPr_element.find(qn("w:numPr"))
    if numPr is None:
        return False
    numId = numPr.find(qn("w:numId"))
    if numId is not None and numId.get(qn("w:val")) == "0":
        return False
    return True


def is_list_item(paragraph) -> bool:
    """判断段落是否为列表项（原生 numPr、列表样式、或纯文本项目符号前缀）。"""
    if has_active_numpr(paragraph._p.pPr):
        return True

    visited = set()
    curr = paragraph.style
    list_keywords = ("list", "bullet", "number", "列表", "项目符号", "编号")
    while curr is not None and curr.name not in visited:
        visited.add(curr.name)
        style_el = getattr(curr, "_element", None)
        if style_el is not None and has_active_numpr(style_el.find(qn("w:pPr"))):
            return True
        if curr.name and any(kw in curr.name.lower() for kw in list_keywords):
            return True
        curr = getattr(curr, "base_style", None)

    stripped = paragraph.text.lstrip()
    if stripped.startswith(LITERAL_BULLET_PREFIXES):
        return True

    return False


def has_math_content(paragraph) -> bool:
    """判断段落中是否包含 Word 公式（OMML 公式块或行内公式）。"""
    for _ in paragraph._p.iter(m_qn("oMathPara")):
        return True
    for _ in paragraph._p.iter(m_qn("oMath")):
        return True
    return False


def clear_positive_indents(paragraph):
    """清除可能将独立公式段落推向中间的多余左缩进/首行缩进。"""
    pPr = paragraph._p.pPr
    if pPr is None:
        return
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return
    for attr in ("w:left", "w:leftChars", "w:firstLine", "w:firstLineChars", "w:start", "w:startChars"):
        val = ind.get(qn(attr))
        if val is not None:
            try:
                if int(val) > 0:
                    del ind.attrib[qn(attr)]
            except ValueError:
                pass


def align_all_omath_para_left(root_el) -> int:
    """将指定 XML 容器内所有公式块 (m:oMathPara) 强制设为左对齐。"""
    count = 0
    for omath_para in root_el.iter(m_qn("oMathPara")):
        pr = omath_para.find(m_qn("oMathParaPr"))
        if pr is None:
            pr = omath_para.makeelement(m_qn("oMathParaPr"))
            omath_para.insert(0, pr)
        jc = pr.find(m_qn("jc"))
        if jc is None:
            jc = pr.makeelement(m_qn("jc"))
            pr.insert(0, jc)
        else:
            jc.attrib.clear()
        jc.set(m_qn("val"), "left")
        count += 1
    return count


def count_omath_nodes(root_el) -> int:
    """统计容器内所有 <m:oMath> 公式节点数。"""
    return sum(1 for _ in root_el.iter(m_qn("oMath")))


def resize_math_fonts(root_el) -> int:
    """检查数学公式内部 (<m:oMath>) 的字号，大于五号(10.5磅)的改为五号。"""
    count = 0
    for omath in root_el.iter(m_qn("oMath")):
        for tag in ("w:sz", "w:szCs"):
            for sz in omath.iter(qn(tag)):
                val = sz.get(qn("w:val"))
                if val and val.isdigit() and int(val) > WUHAO_HALF_PT:
                    sz.set(qn("w:val"), str(WUHAO_HALF_PT))
                    count += 1
    return count


def set_paragraph_spacing(paragraph):
    """将段落设为：行距最小值 11 磅，段前 1 磅，段后 1 磅。"""
    fmt = paragraph.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    fmt.line_spacing = Pt(11)
    fmt.space_before = Pt(1)
    fmt.space_after = Pt(1)

    pPr = paragraph._p.pPr
    if pPr is not None:
        spacing = pPr.find(qn("w:spacing"))
        if spacing is not None:
            for attr in ("w:beforeAutospacing", "w:afterAutospacing"):
                if qn(attr) in spacing.attrib:
                    del spacing.attrib[qn(attr)]


def process_paragraph(paragraph, doc, stats, align_body=True):
    """调整段落间距、对齐方式以及内部所有文本的字号。"""
    set_paragraph_spacing(paragraph)
    stats["paragraphs"] += 1

    if align_body:
        is_list = is_list_item(paragraph)
        has_math = has_math_content(paragraph)
        is_justify = paragraph.alignment in (
            WD_ALIGN_PARAGRAPH.JUSTIFY,
            WD_ALIGN_PARAGRAPH.DISTRIBUTE,
        )

        # 所有列表项、含公式的段落、以及两端对齐段落，统一改为左对齐
        if is_list or has_math or is_justify:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if is_list:
                stats["left_aligned_lists"] += 1
            if has_math:
                stats["left_aligned_math_paras"] += 1
                if not is_list:
                    clear_positive_indents(paragraph)

    for run in paragraph.runs:
        size_pt = get_effective_font_size_pt(run, paragraph, doc)
        if size_pt is not None and size_pt > WUHAO_PT:
            run.font.size = Pt(WUHAO_PT)
            stats["resized_runs"] += 1


def process_styles(doc, stats):
    """将所有字号大于五号(10.5磅)的样式修改为五号，并将两端对齐样式改为左对齐。"""
    for style in doc.styles:
        if hasattr(style, "font") and style.font is not None and style.font.size is not None:
            if style.font.size.pt > WUHAO_PT:
                style.font.size = Pt(WUHAO_PT)
                stats["resized_styles"] += 1

        if hasattr(style, "paragraph_format") and style.paragraph_format is not None:
            if style.paragraph_format.alignment in (
                WD_ALIGN_PARAGRAPH.JUSTIFY,
                WD_ALIGN_PARAGRAPH.DISTRIBUTE,
            ):
                style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_word_field(paragraph, field_name, font_size_pt=WUHAO_PT):
    """向段落插入标准的 Word 域（如 PAGE、NUMPAGES）。"""
    r_begin = paragraph.add_run()
    r_begin.font.size = Pt(font_size_pt)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    r_begin._r.append(fld_begin)

    r_instr = paragraph.add_run()
    r_instr.font.size = Pt(font_size_pt)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_name} "
    r_instr._r.append(instr)

    r_sep = paragraph.add_run()
    r_sep.font.size = Pt(font_size_pt)
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r_sep._r.append(fld_sep)

    r_val = paragraph.add_run("1")
    r_val.font.size = Pt(font_size_pt)

    r_end = paragraph.add_run()
    r_end.font.size = Pt(font_size_pt)
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r_end._r.append(fld_end)


def add_page_number_to_footer(footer):
    """在页脚中间位置添加 '当前页/总页数'（如 1/4）格式页码。"""
    if len(footer.paragraphs) == 1 and not footer.paragraphs[0].text.strip():
        p = footer.paragraphs[0]
    else:
        p = footer.add_paragraph()

    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p)

    add_word_field(p, "PAGE", WUHAO_PT)
    sep_run = p.add_run("/")
    sep_run.font.size = Pt(WUHAO_PT)
    add_word_field(p, "NUMPAGES", WUHAO_PT)


def add_footer_page_numbers(doc):
    """为文档所有节的页脚添加居中 '1/4' 格式页码。"""
    even_odd = doc.settings.odd_and_even_pages_header_footer

    for idx, section in enumerate(doc.sections):
        if idx == 0 or not section.footer.is_linked_to_previous:
            add_page_number_to_footer(section.footer)

        if section.different_first_page_header_footer:
            if idx == 0 or not section.first_page_footer.is_linked_to_previous:
                add_page_number_to_footer(section.first_page_footer)

        if even_odd:
            if idx == 0 or not section.even_page_footer.is_linked_to_previous:
                add_page_number_to_footer(section.even_page_footer)


def format_docx_document(doc: Document, output_path: Path):
    """对已加载的 Document 对象执行全部格式化并保存。"""
    stats = {
        "paragraphs": 0,
        "resized_runs": 0,
        "resized_math_sz": 0,
        "resized_styles": 0,
        "left_aligned_lists": 0,
        "left_aligned_math_paras": 0,
        "omath_nodes": 0,
        "omath_para_blocks": 0,
        "default_changed": False,
        "math_def_changed": False,
    }

    # 1. 遍历文档正文 XML 树中的所有 <w:p>（包含普通段落、表格、内容控件、文本框）
    for p_el in doc.element.body.iter(qn("w:p")):
        paragraph = Paragraph(p_el, doc._body)
        process_paragraph(paragraph, doc, stats, align_body=True)

    # 2. 处理公式对齐与公式字号
    stats["omath_nodes"] = count_omath_nodes(doc.element.body)
    stats["omath_para_blocks"] = align_all_omath_para_left(doc.element.body)
    stats["resized_math_sz"] = resize_math_fonts(doc.element.body)
    stats["math_def_changed"] = set_math_default_left(doc)

    # 3. 处理页眉和页脚中已有段落
    for section in doc.sections:
        headers_footers = [
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ]
        for hf in headers_footers:
            if hf is not None:
                for p_el in hf._element.iter(qn("w:p")):
                    paragraph = Paragraph(p_el, hf)
                    process_paragraph(paragraph, doc, stats, align_body=False)

    # 4. 调整样式定义及 docDefaults 中大于五号的字体
    process_styles(doc, stats)
    stats["default_changed"] = fix_doc_defaults_font_size(doc)

    # 5. 在页脚中间添加 "1/4" 格式页码
    add_footer_page_numbers(doc)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    print(f"处理完成！最终文件已保存至: {output_path}")
    print(
        f"统计信息:\n"
        f"  - 共调整段落: {stats['paragraphs']} 个\n"
        f"  - 设为左对齐的列表项: {stats['left_aligned_lists']} 个，含公式段落: {stats['left_aligned_math_paras']} 个\n"
        f"  - 检测到公式节点 (<m:oMath>): {stats['omath_nodes']} 个，其中独立公式块 (<m:oMathPara>): {stats['omath_para_blocks']} 个（均已设为左对齐）\n"
        f"  - 压缩至五号的文本: {stats['resized_runs']} 处，公式字号: {stats['resized_math_sz']} 处，样式: {stats['resized_styles']} 个"
        f"{'，文档默认字号已同步调整' if stats['default_changed'] else ''}"
        f"{'，全局公式默认对齐已设为left' if stats['math_def_changed'] else ''}\n"
        f"  - 已添加页脚居中页码 (PAGE/NUMPAGES)"
    )


def convert_and_format(input_path: Path, output_path: Path, with_toc: bool = True):
    """自动识别输入类型：Markdown 先转 docx 再排版；docx 直接排版。"""
    if input_path.suffix.lower() == ".docx":
        doc = Document(str(input_path))
        format_docx_document(doc, output_path)
        return

    with tempfile.TemporaryDirectory() as tmp_dir:
        raw_docx_path = Path(tmp_dir) / "raw.docx"
        run_pandoc_to_docx(input_path, raw_docx_path, with_toc=with_toc)
        doc = Document(str(raw_docx_path))
        replace_placeholder_with_tab(doc)
        format_docx_document(doc, output_path)


def main():
    parser = argparse.ArgumentParser(
        description="将 Markdown (.md) 转为 Word (.docx) 并自动完成行距(最小11磅)、段间距(1磅)、字号(<=五号)、列表与公式左对齐及页脚 1/4 页码排版（也支持直接输入 .docx）。"
    )
    parser.add_argument(
        "input_path",
        help="输入文件路径（.md 或 .docx，支持绝对路径或相对路径；若无后缀默认按 .md 处理）",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="输出 .docx 路径（默认：.md 输入生成同名 .docx；.docx 输入生成 原名_adjusted.docx）",
    )
    parser.add_argument(
        "-i",
        "--inplace",
        action="store_true",
        help="当输入为 .docx 时直接覆盖原文件",
    )
    parser.add_argument(
        "--no-toc",
        action="store_true",
        help="转换 Markdown 时不生成目录 (--toc)",
    )

    args = parser.parse_args()
    input_path = Path(args.input_path).expanduser()
    if not input_path.suffix:
        input_path = input_path.with_suffix(".md")
    input_path = input_path.resolve()

    if not input_path.exists() or not input_path.is_file():
        print(f"错误：找不到文件 -> {input_path}")
        sys.exit(1)

    suffix = input_path.suffix.lower()
    if suffix not in (".md", ".markdown", ".docx"):
        print(f"错误：仅支持 .md / .markdown / .docx 文件，当前为: {input_path.name}")
        sys.exit(1)

    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    elif suffix == ".docx":
        output_path = input_path if args.inplace else input_path.with_name(f"{input_path.stem}_adjusted.docx")
    else:
        output_path = input_path.with_suffix(".docx")

    try:
        convert_and_format(input_path, output_path, with_toc=not args.no_toc)
    except subprocess.CalledProcessError as e:
        print(f"Pandoc 执行失败，错误码: {e.returncode}")
        sys.exit(e.returncode)
    except FileNotFoundError as e:
        if "pandoc" in str(e).lower():
            print("错误：未找到 pandoc 命令，请确认已安装 Pandoc 并添加到环境变量 PATH 中。")
        else:
            print(f"错误：{e}")
        sys.exit(1)


if __name__ == "__main__":
    main()