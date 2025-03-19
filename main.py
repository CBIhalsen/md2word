

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
将 Markdown 文件完整转换为 Word (docx) 文件，并能正确识别以下类型的标记:
- 行内公式: $...$、\(...\)
- 块级公式: $$...$$、\[...\]
- Markdown 星号: *斜体*，**加粗**，***粗斜体***
- 表格单元格中的行内公式与星号
- 代码块(```...```)原样写入

在处理行内文本时，我们会统一拆分:
(\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|\*[^*]+\*|\${2}.*?\${2}|\\\[.*?\\\]|\\\(.*?\)|\$.*?\$)

需要:
1) pip install latex2mathml python-docx lxml markdown
2) 在脚本同目录下放 MML2OMML.XSL (或自行修改 xsl_path).
"""

import os
import re
import markdown
import latex2mathml.converter
from latex2mathml.exceptions import NoAvailableTokensError
from lxml import etree
from docx import Document
from docx import shared
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH  # 新增
# --------------------- 新增部分：用于插入水平线 -----------------------------
def insert_horizontal_rule(doc):
    """
    在 doc 中插入一个“水平分割线”段落 (通过底部边框模拟)。
    """
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert(0, pBdr)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')     # 线条样式
    bottom.set(qn('w:sz'), '6')          # 线条粗细
    bottom.set(qn('w:space'), '1')       # 与文字的间距
    bottom.set(qn('w:color'), 'auto')    # 颜色
    pBdr.append(bottom)
# --------------------- 新增部分结束 ----------------------------------------

def latex_to_omml(latex_input, xsl_path):
    """
    将单个 LaTeX 公式字符串转换为可供 python-docx 使用的 OMML XML 元素。
    带调试信息，查看传入与解析结果。
    """
    print(f"\n[DEBUG] -> latex_to_omml收到公式: {repr(latex_input)}")
    latex_input_stripped = latex_input.strip()
    if not latex_input_stripped:
        print("[DEBUG] -> 公式内容为空，返回 None")
        return None
    try:
        # 若 latex2mathml 对某些符号(例\nabla)不支持，可在此做替换
        # latex_input_stripped = latex_input_stripped.replace(r'\nabla', '∇')
        mathml_output = latex2mathml.converter.convert(latex_input_stripped)
        print(f"[DEBUG] -> latex2mathml转换成功, MathML片段(截断): {mathml_output[:80]}...")
        xsl_tree = etree.parse(xsl_path)
        transform = etree.XSLT(xsl_tree)
        mathml_dom = etree.fromstring(mathml_output)
        omml_dom = transform(mathml_dom)
        print(f"[DEBUG] -> XSLT转换成功, OMML(截断): {etree.tostring(omml_dom)[:80]}...")
        return omml_dom.getroot()
    except (NoAvailableTokensError, etree.XMLSyntaxError) as e:
        print(f"[DEBUG] -> latex2mathml或XML解析异常: {e}")
        return None
    except Exception as e:
        print(f"[DEBUG] -> 其他异常: {e}")
        return None

def strip_delimiters(latex_str):
    """
    去掉行内或块级分隔符 ($, $$, \(, \[ 等) 后得到纯 LaTeX 公式内容。
    带调试信息。
    """
    print(f" [DEBUG] strip_delimiters原始: {repr(latex_str)}")
    pattern = r'^(?:\${1,2}|\\\(|\\\[)\s*|\s*(?:\${1,2}|\\\)|\\\])$'
    result = re.sub(pattern, '', latex_str.strip())
    print(f" [DEBUG] strip_delimiters后: {repr(result)}")
    return result.strip()

def add_runs_with_inline_markdown(text, paragraph, xsl_path):
    """
    拆分行内文本, 同时支持:
    - Markdown星号: *...*, **...**, ***...***
    - LaTeX公式: $...$, $$...$$, \(...\), \[...\]
    并在python-docx的paragraph中插入对应的Run或OMML节点。
    """
    print(f"[DEBUG] -> 解析行内标记: {repr(text)}")
    # 修正后的正则表达式:
    inline_pattern = (
        r'(\*\*\*[^*]+\*\*\*|'    # ***粗斜体***
        r'\*\*[^*]+\*\*|'         # **加粗**
        r'\*[^*]+\*|'             # *斜体*
        r'\${2}[\s\S]*?\${2}|'    # $$公式$$
        r'\\\[[\s\S]*?\\\]|'      # \[公式\]
        r'\\\([\s\S]*?\\\)|'      # \(公式\)
        r'\$[\s\S]*?\$)'          # $公式$
    )
    segments = re.split(inline_pattern, text, flags=re.DOTALL)
    print(f"[DEBUG] -> re.split切分片段: {segments}")

    for seg in segments:
        # 如果 seg 能二次匹配说明它是一个公式或星号标记
        if re.match(inline_pattern, seg, flags=re.DOTALL):
            # (A) 先判断是否是公式
            formula_pattern = r'(\${1,2}[\s\S]*?\${1,2}|\\\[[\s\S]*?\\\]|\\\([\s\S]*?\\\))'
            if re.match(formula_pattern, seg, flags=re.DOTALL):
                print(f"[DEBUG] -> 检测到公式段: {repr(seg)}")
                latex_code_inline = strip_delimiters(seg)
                omml_elem = latex_to_omml(latex_code_inline, xsl_path)
                if omml_elem is not None:
                    paragraph._element.append(omml_elem)
                else:
                    paragraph.add_run(seg)  # 转换失败则原样
            else:
                # (B) 否则就是星号标记
                print(f"[DEBUG] -> 检测到星号标记: {repr(seg)}")
                triple_star_pattern = r'^\*\*\*([^*]+)\*\*\*$'
                double_star_pattern = r'^\*\*([^*]+)\*\*$'
                single_star_pattern = r'^\*([^*]+)\*$'
                m3 = re.match(triple_star_pattern, seg)
                m2 = re.match(double_star_pattern, seg)
                m1 = re.match(single_star_pattern, seg)
                if m3:
                    content = m3.group(1)
                    run_obj = paragraph.add_run(content)
                    run_obj.bold = True
                    run_obj.italic = True
                elif m2:
                    content = m2.group(1)
                    run_obj = paragraph.add_run(content)
                    run_obj.bold = True
                elif m1:
                    content = m1.group(1)
                    run_obj = paragraph.add_run(content)
                    run_obj.italic = True
                else:
                    paragraph.add_run(seg)
        else:
            # 普通文本
            paragraph.add_run(seg)

def convert_markdown_to_docx(md_file, xsl_path, output_docx):
    """
    将 Markdown 文件转换为 Word (docx).
    - 行内星号(斜体/加粗/粗斜体) 和 行内公式(LaTeX) 都能被识别
    - 多行块公式、表格、代码块等也保留
    """
    print(f"[DEBUG] -> 开始读取Markdown文件: {md_file}")
    with open(md_file, 'r', encoding='utf-8') as f:
        md_text = f.read()
    print("[DEBUG] -> 读取完成, 字符数: ", len(md_text))

    # 测试下 markdown -> html, 不做实际处理
    _ = markdown.markdown(md_text, extensions=["tables", "fenced_code"])

    doc = Document()
    doc.styles['Normal'].font.name = 'Times New Roman'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    doc.styles['Normal'].font.size = shared.Pt(10)

    lines = md_text.split('\n')

    in_code_block = False
    in_math_block = False
    math_block_buffer = []

    # 表格相关
    table_buffer = []
    table_mode = False

    def flush_table(doc_obj):
        """
        将 table_buffer 中的行转换成表格插入 docObj，并清空。
        在插入过程中解析单元格中的行内星号和公式。
        """
        nonlocal table_buffer
        if not table_buffer:
            return

        print(f"[DEBUG] -> flush_table, 行数: {len(table_buffer)}")
        rows_data = []
        for row in table_buffer:
            row = row.strip().strip('|')
            row_cells = [cell.strip() for cell in row.split('|')]
            rows_data.append(row_cells)

        nrows = len(rows_data)
        ncols = len(rows_data[0]) if rows_data else 0
        if nrows == 0 or ncols == 0:
            table_buffer = []
            return

        table = doc_obj.add_table(rows=nrows, cols=ncols)
        table.style = 'Table Grid'
        for i in range(nrows):
            for j in range(ncols):
                cell_paragraph = table.cell(i, j).paragraphs[0]
                cell_paragraph.text = ''
                cell_text = rows_data[i][j]
                add_runs_with_inline_markdown(cell_text, cell_paragraph, xsl_path)

        table_buffer = []

    print("[DEBUG] -> 开始逐行处理Markdown...")
    for line in lines:
        line_stripped = line.strip()
        print(f"\n[DEBUG] -> 处理行: {repr(line_stripped)}")

        # 代码块检测
        if line_stripped.startswith('```'):
            in_code_block = not in_code_block
            print(f"[DEBUG] -> {'进入' if in_code_block else '退出'}代码块")
            if in_code_block:
                doc.add_paragraph("----- 代码块开始 -----")
            else:
                doc.add_paragraph("----- 代码块结束 -----")
            continue

        if in_code_block:
            # 代码块区域，原样写入
            doc.add_paragraph(line)
            continue

        # 若在多行块公式中
        if in_math_block:
            if line_stripped.endswith('$$') or line_stripped.endswith('\\]'):
                math_block_buffer.append(line)
                print("[DEBUG] -> 块公式结束, 开始转换")
                in_math_block = False
                block_content = "\n".join(math_block_buffer)
                block_content_stripped = strip_delimiters(block_content)
                paragraph = doc.add_paragraph()
                omml_elem = latex_to_omml(block_content_stripped, xsl_path)
                if omml_elem is not None:
                    paragraph._element.append(omml_elem)
                else:
                    paragraph.add_run(block_content)
                math_block_buffer = []
                continue
            else:
                math_block_buffer.append(line)
                continue

        # 检查多行块公式的开始
        if (
            (line_stripped.startswith('$$') or line_stripped.startswith('\\[')) and
            not (line_stripped.endswith('$$') or line_stripped.endswith('\\]'))
        ):
            print("[DEBUG] -> 多行块公式开始")
            in_math_block = True
            math_block_buffer = [line]
            continue

        # 单行块公式
        block_patt = r'^(?:\${2}.*?\${2}|\\\[.*?\\\])$'
        if re.match(block_patt, line_stripped):
            print("[DEBUG] -> 单行块公式")
            latex_code_block = strip_delimiters(line_stripped)
            paragraph = doc.add_paragraph()
            omml_elem = latex_to_omml(latex_code_block, xsl_path)
            if omml_elem is not None:
                paragraph._element.append(omml_elem)
            else:
                paragraph.add_run(line_stripped)
            continue

        # 空行
        if not line_stripped:
            print("[DEBUG] -> 空行")
            if table_mode:
                flush_table(doc)
                table_mode = False
            doc.add_paragraph("")
            continue

        # 表格行
        if line_stripped.startswith('|') and line_stripped.endswith('|'):
            # 如果整行只含 '-', '|', ':' 等字符(典型 markdown 表头分隔)，跳过
            if all(ch in '-|: ' for ch in line_stripped):
                print("[DEBUG] -> 检测到 Markdown 表格分隔行，跳过")
                continue
            print("[DEBUG] -> 表格行")
            table_buffer.append(line_stripped)
            table_mode = True
            continue
        else:
            if table_mode:
                flush_table(doc)
                table_mode = False

        # --------------------- 新增逻辑：检测 '---' 就插入水平线 ---------------------
        if line_stripped == '---':
            print("[DEBUG] -> 检测到水平线标记 '---'")
            insert_horizontal_rule(doc)   # 调用我们上面新增的函数
            continue
        # --------------------- 新增逻辑结束 -----------------------------------------

        # 标题(# 开头)
        if line_stripped.startswith('#'):
            match_hashes = re.match(r'^(#+)', line_stripped)
            level = len(match_hashes.group(1)) if match_hashes else 1
            heading_text = line_stripped[level:].strip()
            print(f"[DEBUG] -> 标题: level={level}, text={heading_text}")
            doc.add_heading(heading_text, min(level, 6))
            continue

        # 无序列表
        if line_stripped.startswith('* ') or line_stripped.startswith('- '):
            item_text = line_stripped[1:].strip()
            print(f"[DEBUG] -> 无序列表项: {item_text}")
            paragraph = doc.add_paragraph(style=None)
            paragraph.add_run("• " + item_text)
            continue

        # 剩下的普通行(含可能的行内公式/星号)
        print("[DEBUG] -> 普通文本行(可能含行内公式或星号)")
        paragraph = doc.add_paragraph(style=None)
        add_runs_with_inline_markdown(line_stripped, paragraph, xsl_path)

    # 结尾可能还有表格未flush
    if table_mode:
        flush_table(doc)

    doc.save(output_docx)
    print(f"[DEBUG] -> 已生成 Word 文档: {output_docx}")

if __name__ == "__main__":
    # 设置 XSLT路径
    xsl_path = os.path.join('.', 'MML2OMML.XSL')
    md_file = "example.md"  # 你的Markdown文件
    output_docx = "output_from_markdown_stars_and_formula.docx"
    convert_markdown_to_docx(md_file, xsl_path, output_docx)
