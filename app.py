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
from flask import abort

matplotlib.use('Agg')  # 设置为非交互式后端

from openpyxl.reader.excel import load_workbook

app = Flask(__name__)

matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 黑体
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 文字保存路径
SAVE_DIR = "saved_texts"

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


# @app.route('/')
# def home():
#     return render_template('version1.html')


@app.route('/api/problem-lines')
def api_problem_lines():
    problem_df = process_excel_data()
    result = problem_df.to_dict('records')
    return jsonify({
        'success': True,
        'data': result,
        'total': len(result)
    })


@app.route('/<page_id>')
def dynamic_page(page_id):
    """根据URL自动匹配对应模板"""
    template_path = os.path.join(app.template_folder, f"{page_id}.html")
    if os.path.exists(template_path):
        return render_template(f"{page_id}.html")
    else:
        return render_template("not_found.html", page_id=page_id), 404


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
                for side in ['top', 'bottom', 'left', 'right']:
                    mar = docx.oxml.shared.OxmlElement(f'w:{side}')
                    mar.set(docx.oxml.shared.qn('w:w'), '50')
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
        plt.pie([heavy, light], labels=['重载', '轻载'], autopct='%1.1f%%', colors=['#dc3545', '#198754'])
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


# 在文件顶部添加Excel文件路径
EXCEL_FILE = 'data/收资清单填充结果1.xlsx'

# ---定义完整的表格配置字典 ---
# 这是整个系统的核心配置，请务必根据您的实际文档和Excel文件仔细填写
# TABLE_CONFIG = {
#     'table-2-1': {
#         'title': '表2-1 xx网格35kV及以上变电站基本情况表',
#         'sheet_name': '高压变电站基础信息表',  # 对应Excel中的Sheet名
#         'columns': [
#             {'header': '序号', 'source': '序号'},
#             {'header': '变电站名称', 'source': '变电站名称'},
#             {'header': '总容量', 'source': '总容量（MVA）'},
#             {'header': '主变编号', 'source': '主变编号'},
#             {'header': '主变容量', 'source': '主变容量（MVA）'},
#             # 针对“10kV馈线间隔”的精确映射，一个表头对应一个源列
#             {'header': '10kV馈线间隔总数', 'source': '10kV馈线间隔总数'},
#             {'header': '10kV馈线间隔已使用', 'source': '10kV馈线间隔已使用'},
#             {'header': '10kV馈线间隔剩余', 'source': '10kV馈线间隔剩余'},
#             {'header': '未来可扩展间隔数', 'source': '未来可扩展间隔数（个）'},
#             {'header': '备注', 'source': '备注'},
#         ]
#     },
#     'table-2-2': {
#         'title': '表2-2 xx网格35kV及以上变电站负载情况表',
#         'sheet_name': '高压变电站基础信息表',  # 同一个Sheet可以供给多个表格
#         'columns': [
#             {'header': '编号', 'source': '序号'},  # '序号'列在文档中可能被称为'编号'
#             {'header': '变电站名称', 'source': '变电站名称'},
#             {'header': '总容量', 'source': '总容量（MVA）'},
#             {'header': '主变编号', 'source': '主变编号'},
#             {'header': '主变容量', 'source': '主变容量（MVA）'},
#             {'header': '主变典型日负荷', 'source': '主变典型日负荷（MW）'},
#             {'header': '主变典型日负载率', 'source': '主变典型日负载率（%）'},
#             {'header': '主变年最大负荷', 'source': '主变年最大负荷（MW）'},
#             {'header': '主变年最大负载率', 'source': '主变年最大负载率（%）'},
#         ]
#     },
#     'table-2-3': {
#         'title': '表2-3 xx网格供电区域概况表',
#         'sheet_name': '供电区域概况表',
#         'columns': [
#             {'header': '网格', 'source': '网格名称'},
#             {'header': '供区类型', 'source': '供区类型'},
#             {'header': '指标分项', 'source': '指标'},
#             {'header': '数值', 'source': '数值'},
#             {'header': '单位', 'source': '单位'},
#         ]
#     },
#     'table-3-1': {
#         'title': '表3-1 xx网格10kV线路基本情况表',
#         'sheet_name': '10kV线路基础信息表',
#         'columns': [
#             {'header': '线路名称', 'source': '线路名称'},
#             {'header': '起点站', 'source': '起点变电站'},
#             {'header': '联络点', 'source': '主要联络点'},
#             {'header': '线路型号', 'source': '导线型号'},
#             {'header': '线路长度', 'source': '线路长度'},
#             {'header': '开关状态', 'source': '开关状态'},
#         ]
#     },
#     # ... 在此继续添加其他表格的配置
# }
#
# 3333
# # ---  创建新的数据获取路由 ---
# @app.route('/get_table_data/<table_id>')
# def get_table_data(table_id):
#     """
#     根据table_id从配置中获取数据，并返回JSON格式
#     """
#     config = TABLE_CONFIG.get(table_id)
#     if not config:
#         return jsonify({'error': 'Table configuration not found'}), 404
#
#     try:
#         # 1. 读取指定的Excel Sheet
#         df = pd.read_excel(EXCEL_FILE, sheet_name=config['sheet_name'])
#
#         # 2. 提取需要的源列
#         source_columns = [col['source'] for col in config['columns']]
#
#         # 检查所有需要的列是否存在
#         missing_cols = [col for col in source_columns if col not in df.columns]
#         if missing_cols:
#             return jsonify({'error': f'Columns not found in Excel sheet: {", ".join(missing_cols)}'}), 400
#
#         filtered_df = df[source_columns]
#
#         # 3. 重命名列，使其与HTML表头一致
#         # 创建一个从'source'到'header'的映射字典
#         rename_map = {col['source']: col['header'] for col in config['columns']}
#         filtered_df = filtered_df.rename(columns=rename_map)
#
#         # 4. 处理NaN值，避免前端显示问题
#         filtered_df = filtered_df.fillna('')
#
#         # 5. 转换为JSON格式
#         data = filtered_df.to_dict(orient='records')
#         headers = [col['header'] for col in config['columns']]
#
#         return jsonify({
#             'title': config['title'],
#             'headers': headers,
#             'data': data
#         })
#
#     except FileNotFoundError:
#         return jsonify({'error': f'Excel file not found at {EXCEL_FILE}'}), 500
#     except Exception as e:
#         return jsonify({'error': str(e)}), 500
#
#
# 22222
# 添加sheet名称映射字典
SHEET_MAPPING = {
    'table2-1': '1高压变电站基础信息表',
    'table2-2': '2供电区域概况表',
    'table2-3': '3配变负载率明细表',
    'table2-4': '4中压线路负载率表',
    'table2-5': '5高层小区供电保障表',
    'table2-6': '6分布式光伏接入情况表',
    'table2-7': '7充电设施统计表',
    'table2-8': '8台区停电明细表',
    'table2-9': '9台区电压越限统计表',
    'table2-10': '10中压线路网架表',
    'table2-11': '11中压线路分段情况',
    'table2-12': '12中压线路 N-1 校验表',
    'table2-13': '13开关明细表',
    'table2-14': '14配变明细表',
    'table2-15': '15线路规模',
    'table2-16': '16 0.4kV 台区供电半径表',
    'table2-17': '17 10kV线路投运时间表',
    'table2-18': '18线路分线线损情况明细表',
    'table2-19': '19台区线损明细',
    'table2-20': '20 10kV线路配自终端明细表',
    'table2-21': '21线路跨网格供电情况',
    'table2-22': '22线路装接配变容量',
    'table2-23': '23配变情况分布表',
    'table2-24': '24干线型号统计表',
    'table2-25': '25环网柜明细表',
    'table2-26': '26中压线路大分支情况',
    'table2-27': '27中压线路供电半径',
    'table2-28': '28线路状态检测',
    'table2-29': '29问题清单汇总',
    # 添加其他需要的映射...
}


# 添加通用表格路由
# 在app.py顶部添加表格配置字典
TABLE_CONFIG = {
    '2-1': {
        'sheet_name': '1高压变电站基础信息表',
        'columns': ['序号', '变电站名称', '总容量（MVA）', '主变编号', '主变容量（MVA）', '10kV馈线间隔（个）', '未来可扩展间隔数（个）', '备注'],
        'title': '表2-1 xx网格35kV及以上变电站基本情况表'
    },
    '2-2': {
        'sheet_name': '1高压变电站基础信息表',
        'columns': ['编号', '变电站名称', '总容量（MVA）', '主变编号', '主变容量（MVA）', '主变典型日负荷（MW）', '主变典型日负载率（%）', '主变年最大负荷（MW）', '主变年最大负载率（%）'],
        'title': '表2-2 xx网格35kV及以上变电站基本情况表'
    },
    '2-3': {
        'sheet_name': '2供电区域概况表',
        'columns': ['网格', '供区类型', '指标分项', '数值'],
        'title': '表2-3 低压电网规模一览表'
    },
    # 添加其他表格配置...
}

# 修改通用表格路由
@app.route('/table<path:table_id>')
def show_table(table_id):
    config = TABLE_CONFIG.get(table_id)
    if not config:
        return "未找到对应的表格", 404

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=config['sheet_name'])
        # 只保留配置中指定的列
        existing_columns = [col for col in config['columns'] if col in df.columns]
        df = df[existing_columns]
        df = df.fillna('')
        table_html = df.to_html(classes='table table-striped table-hover', index=False, table_id='data-table')
        return render_template('table_template.html', table_html=table_html, title=config['title'])
    except Exception as e:
        return f"读取表格出错: {str(e)}", 500

# 1111111

@app.route('/api/sheets')
def get_sheets():
    """获取Excel文件中所有sheet的名称"""
    try:
        xls = pd.ExcelFile('data/收资清单填充结果1.xlsx')
        sheet_names = xls.sheet_names
        return jsonify(sheet_names)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/sheet/<sheet_name>')
def get_sheet_data(sheet_name):
    """获取指定sheet的数据"""
    try:
        # 读取指定sheet的数据
        df = pd.read_excel('data/收资清单填充结果1.xlsx', sheet_name=sheet_name)
        # 处理NaN值
        df = df.fillna('')
        # 转换为JSON格式
        data = df.to_dict('records')
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 添加新路由
@app.route('/all_sheets')
def all_sheets():
    """展示所有sheet的页面"""
    return render_template('all_sheets.html')
@app.route('/version1')
def version1():
    """展示所有sheet的页面"""
    return render_template('version1.html')

if __name__ == '__main__':
    app.run(debug=True, port=5001)
