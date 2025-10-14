import os
import docx
from flask import Flask, render_template, jsonify, request, send_file
import pandas as pd
from io import BytesIO
from docx import Document
import matplotlib
import matplotlib.pyplot as plt
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import matplotlib.pyplot as plt
from io import BytesIO
import pandas as pd
from docx.shared import Pt, Inches
matplotlib.use('Agg')  # 设置为非交互式后端

from openpyxl.reader.excel import load_workbook

app = Flask(__name__)

matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 黑体
matplotlib.rcParams['axes.unicode_minus'] = False    # 解决负号显示问题

# 文字保存路径
SAVE_DIR = "saved_pages"

os.makedirs(SAVE_DIR, exist_ok=True)

def get_save_path(page_id):
    return os.path.join(SAVE_DIR, f"{page_id}.txt")


def read_excel(file_path, sheet_name='Sheet1'):
    """
    直接读取Excel文件，不需要处理合并单元格
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df


def process_excel_data():
    """读取并处理Excel数据，拆分合并单元格，只返回重载或轻载线路"""
    try:
        df = read_excel('test.xlsx', sheet_name='Sheet1')

        # 问题线路条件：重载≥90%，轻载≤25%
        problem_conditions = (
            (df['2024年线路最大负载率（%）'] >= 90) |
            (df['2024年线路最大负载率（%）'] <= 25)
        )

        problem_df = df[problem_conditions].copy()
        problem_df = problem_df.sort_values('2024年线路最大负载率（%）', ascending=False)
        problem_df = problem_df.reset_index(drop=True)
        problem_df['序号'] = problem_df.index + 1

        # 空值填充只针对非“重载/轻载原因”列
        problem_df = problem_df.fillna('')

        return problem_df


    except Exception as e:
        print(f"读取文件错误: {e}")
        return pd.DataFrame()


@app.route('/api/problem-lines')  
def api_problem_lines():
    problem_df = process_excel_data()
    result = problem_df.to_dict('records')
    return jsonify({
        'success': True,
        'data': result,
        'total': len(result)
    })

@app.route('/line-load')
def line_load():
    return render_template('2-5.html')

@app.route('/export/<page_id>')
def export_word(page_id):
    """导出Word文件，包含表格、图表和页面文本内容"""
    # 1. 处理Excel数据生成表格和图表
    df = process_excel_data()
    
    # 2. 读取页面文本内容
    save_path = get_save_path(page_id)
    try:
        with open(save_path, "r", encoding="utf-8") as f:
            page_text = f.read()
    except FileNotFoundError:
        page_text = "暂无内容。"

    doc = Document()
    
    # 第一部分：页面文本内容
    doc.add_heading(f"页面 {page_id} 内容", level=1)
    doc.add_paragraph(page_text)
    
    # 添加分页符（可选）
    # doc.add_page_break()
    
    # 第二部分：问题线路分析报告（仅当有数据时）
    if not df.empty:
        doc.add_heading("问题线路分析报告", level=1)

        # 插入表格
        table = doc.add_table(rows=1, cols=len(df.columns))
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(df.columns):
            hdr_cells[i].text = str(col)
            # 设置表头字体大小
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)

        # 设置单元格内边距更紧凑
        table.allow_autofit = True
        for row in table.rows:
            for cell in row.cells:
                cell.width = Inches(1.0)
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcMar = docx.oxml.shared.OxmlElement('w:tcMar')
                for side in ['top','bottom','left','right']:
                    mar = docx.oxml.shared.OxmlElement(f'w:{side}')
                    mar.set(docx.oxml.shared.qn('w:w'),'50')
                    tcMar.append(mar)
                tcPr.append(tcMar)

        # 填充数据
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, val in enumerate(row):
                row_cells[i].text = str(val)
                for paragraph in row_cells[i].paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)

        # 第三部分：图表
        import matplotlib.pyplot as plt
        from io import BytesIO

        # 柱状图
        plt.figure(figsize=(10, 6))
        top10 = df.nlargest(10, '2024年线路最大负载率（%）')
        plt.bar(top10['线路名称'], top10['2024年线路最大负载率（%）'], color='#0d6efd')
        plt.xticks(rotation=45, ha='right')
        plt.ylabel("负载率（%）")
        plt.title("前10条线路负载率情况")
        buf1 = BytesIO()
        plt.tight_layout()
        plt.savefig(buf1, format="png", dpi=150)
        plt.close()
        buf1.seek(0)
        doc.add_picture(buf1, width=Inches(5.5))

        # 饼图
        plt.figure(figsize=(8, 6))
        heavy = (df['2024年线路最大负载率（%）'] >= 90).sum()
        light = (df['2024年线路最大负载率（%）'] <= 25).sum()
        plt.pie([heavy, light], labels=['重载','轻载'], autopct='%1.1f%%', colors=['#dc3545','#198754'])
        plt.title("重载/轻载占比")
        buf2 = BytesIO()
        plt.savefig(buf2, format="png", dpi=150)
        plt.close()
        buf2.seek(0)
        doc.add_picture(buf2, width=Inches(5.5))
    else:
        # 如果没有数据，添加提示
        doc.add_heading("数据分析", level=1)
        doc.add_paragraph("暂无数据分析内容。")

    # 保存到内存流
    out_stream = BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    
    return send_file(out_stream, as_attachment=True, download_name=f"{page_id}_完整报告.docx")

@app.route('/page/<page_id>')
def page(page_id):
    """加载指定页面的HTML"""
    # 读取页面文本内容
    save_path = get_save_path(page_id)
    try:
        with open(save_path, "r", encoding="utf-8") as f:
            text = f.read()
    except FileNotFoundError:
        text = "暂无内容，可以点击编辑并保存。"
    
    # 直接渲染对应的HTML模板
    return render_template(f"{page_id}.html", text=text, page_id=page_id)


@app.route('/save/<page_id>', methods=['POST'])
def save_text(page_id):
    """保存指定页面的文本"""
    text = request.json.get("text", "")
    with open(get_save_path(page_id), "w", encoding="utf-8") as f:
        f.write(text)
    return jsonify({"status": "ok", "message": f"页面 {page_id} 保存成功！"})

    
@app.route('/')
def index():
    """首页，列出所有页面链接"""
    pages = ['2-5', '3-3']
    return render_template('index.html', pages=pages)


if __name__ == '__main__':
    app.run(debug=True, port=5001)
