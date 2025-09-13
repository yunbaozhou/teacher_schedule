import pandas as pd
from docx import Document
from docx.shared import Inches, RGBColor, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import json
import random
from flask import Flask, request, jsonify, send_file, render_template
import os
# 添加必要的导入
from io import BytesIO

app = Flask(__name__, static_folder='static', template_folder='templates')

# 颜色映射规则（基于需求分析报告）
COLOR_MAP = {
    "语文": (255, 204, 204),    # 粉色
    "数学": (204, 255, 255),    # 浅蓝色
    "英语": (204, 255, 204),    # 绿色
    "综研": (229, 229, 204),    # 浅绿色
    "趣味体育": (255, 255, 153), # 黄色
    "体育": (153, 204, 255),    # 蓝色
    "音乐": (221, 170, 221),    # 浅紫色
    "体育与健康": (153, 204, 255), # 浅蓝色
    "道法": (204, 153, 204),    # 紫色
    "美术": (255, 179, 136),    # 橙色
    "科学": (255, 229, 153),    # 浅黄色
    "劳动": (204, 255, 204)     # 绿色
}

# 额外的默认颜色，用于自动分配给未定义颜色的课程
DEFAULT_COLORS = [
    (255, 192, 203),  # 粉色
    (173, 216, 230),  # 浅蓝
    (144, 238, 144),  # 浅绿
    (255, 182, 193),  # 浅粉红
    (221, 160, 221),  # 梅花色
    (175, 238, 238),  # 浅青色
    (255, 218, 185),  # 桃色
    (240, 230, 140),  # 卡其色
    (230, 230, 250),  # 薰衣草色
    (255, 228, 196),  # 比卡迪色
    (255, 160, 122),  # 浅珊瑚色
    (176, 224, 230),  # 粉蓝
    (255, 228, 181),  # 麦色
    (189, 183, 107),  # 暗卡其色
    (216, 191, 216),  # 苍紫罗兰色
    (152, 251, 152),  # 薄荷奶油色
    (173, 216, 230),  # 天蓝色
    (255, 192, 203),  # 粉红色
    (244, 164, 96),   # 沙棕色
    (210, 180, 140),  # 萨摩色
    (255, 215, 0),    # 金色
    (218, 112, 214),  # 兰花色
    (192, 192, 192),  # 灰色
    (128, 128, 0),    # 橄榄色
    (128, 0, 128),    # 紫色
    (0, 128, 128),    # 水鸭色
    (0, 0, 128),      # 海军蓝
    (139, 0, 0),      # 深红色
    (0, 100, 0),      # 深绿色
    (128, 0, 0),      # 栗色
]

def get_course_color(course_name, user_selected_colors=None):
    """
    获取课程颜色，优先级：
    1. 用户选择的颜色
    2. 预定义的颜色映射
    3. 自动分配默认颜色（基于课程名称严格计算）
    """
    # 如果用户选择了颜色，则优先使用用户选择的颜色
    if user_selected_colors and course_name in user_selected_colors:
        color = user_selected_colors[course_name]
        # 如果颜色是字符串格式，转换为元组
        if isinstance(color, str):
            return tuple(map(int, color.split(',')))
        return color
    
    # 如果预定义了颜色，则使用预定义颜色
    if course_name in COLOR_MAP:
        return COLOR_MAP[course_name]
    
    # 基于课程名称生成稳定的哈希值，确保相同名称的课程总是获得相同的颜色
    # 使用更大的颜色池以减少不同名称课程获得相同颜色的概率
    hash_value = hash(course_name) % (len(DEFAULT_COLORS) * 100)  # 扩大颜色空间
    color_index = hash_value % len(DEFAULT_COLORS)
    return DEFAULT_COLORS[color_index]

# 用于存储课程数据的全局变量
courses_data_store = []

# 根路由，返回前端页面
@app.route('/')
def index():
    return render_template('schedule.html')

# API路由

@app.route('/api/courses', methods=['GET'])
def get_courses():
    """获取所有课程"""
    return jsonify(courses_data_store)

@app.route('/api/courses', methods=['POST'])
def add_course():
    """添加课程"""
    try:
        course_data = request.json
        courses_data_store.append(course_data)
        return jsonify({"success": True, "message": "课程添加成功"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400


@app.route('/api/courses/conflicts', methods=['POST'])
def check_conflicts():
    """检查课程冲突"""
    try:
        from io import StringIO
        import sys
        
        # 保存原始的stdout
        old_stdout = sys.stdout
        sys.stdout = StringIO()
        
        data = request.json
        courses = data.get('courses', [])
        df = pd.DataFrame(courses)
        conflicts = detect_conflicts(df)
        
        # 恢复原始的stdout
        sys.stdout = old_stdout
        
        return jsonify({"conflicts": conflicts})
    except Exception as e:
        return jsonify({"conflicts": [], "error": str(e)}), 400

@app.route('/api/export/excel', methods=['POST'])
def export_excel():
    """导出为Excel"""
    try:
        data = request.json
        courses = data.get('courses', [])
        title = data.get('title', '课程表')  # 从请求中获取标题
        df = pd.DataFrame(courses)
        
        # 生成Excel文件到内存
        output = BytesIO()
        generate_excel(df, output, title)
        output.seek(0)
        
        # 直接返回文件流
        filename = f"{title}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400

@app.route('/api/export/word', methods=['POST'])
def export_word():
    """导出为Word"""
    try:
        data = request.json
        courses = data.get('courses', [])
        user_selected_colors = data.get('userSelectedColors', {})
        title = data.get('title', '课程表')  # 从请求中获取标题
        df = pd.DataFrame(courses)
        
        # 生成Word文件到内存
        output = BytesIO()
        generate_word(df, output, user_selected_colors, title)
        output.seek(0)
        
        # 直接返回文件流
        filename = f"{title}.docx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400

def detect_conflicts(course_data):
    """检测课程冲突：同一教师在同一时间的课程安排"""
    conflicts = []
    # 按教师和时间分组检查
    if "教师" in course_data.columns and "星期" in course_data.columns and "节次" in course_data.columns:
        grouped = course_data.groupby(["教师", "星期", "节次"])
        for key, group in grouped:
            if len(group) > 1:
                conflicts.append(f"冲突：教师{key[0]}在{key[1]}{key[2]}节有{len(group)}门课程")
    
    # 检查同一班级在同一时间的课程安排
    if "班级" in course_data.columns and "星期" in course_data.columns and "节次" in course_data.columns:
        grouped = course_data.groupby(["班级", "星期", "节次"])
        for key, group in grouped:
            if len(group) > 1:
                conflicts.append(f"冲突：班级{key[0]}在{key[1]}{key[2]}节有{len(group)}门课程")
    
    return conflicts

def generate_excel(df, output, title):
    """生成Excel文件"""
    # 创建工作簿和工作表
    wb = Workbook()
    ws = wb.active
    ws.title = title
    
    # 设置标题
    ws.merge_cells('A1:H1')
    title_cell = ws['A1']
    title_cell.value = title
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 设置表头
    headers = ['节次/星期', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    
    # 设置上午/下午/晚自习标识
    time_periods = ['上午', '下午', '晚自习']
    period_rows = [3, 7, 11]  # 上午从第3行开始，下午从第7行开始，晚自习从第11行开始
    
    for i, period in enumerate(time_periods):
        ws.merge_cells(start_row=period_rows[i], start_column=1, end_row=period_rows[i], end_column=8)
        period_cell = ws.cell(row=period_rows[i], column=1, value=period)
        period_cell.font = Font(bold=True)
        period_cell.alignment = Alignment(horizontal='center', vertical='center')
        period_cell.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    
    # 填充课程数据
    # 创建一个字典来快速查找课程
    course_dict = {}
    for _, course in df.iterrows():
        key = (course['星期'], course['节次'])
        if key not in course_dict:
            course_dict[key] = []
        course_dict[key].append(course)
    
    # 填充课程数据
    for period_rows_idx, start_row in enumerate([4, 8, 12]):  # 课程从第4、7、11行开始
        for row_offset in range(4):  # 每个时段4节课
            row = start_row + row_offset
            # 设置节次
            period = row_offset + 1 + period_rows_idx * 4
            ws.cell(row=row, column=1, value=f'第{period}节')
            
            # 填充每天的课程
            days = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
            for col, day in enumerate(days, 2):
                key = (day, period)
                if key in course_dict:
                    courses = course_dict[key]
                    # 如果同一时间有多门课程，显示所有课程
                    course_texts = []
                    for course in courses:
                        course_text = course['课程名称']
                        if course.get('教师') and course['教师'] != '未指定':
                            course_text += f"\n教师：{course['教师']}"
                        if course.get('地点') and course['地点'] != '未指定':
                            course_text += f"\n地点：{course['地点']}"
                        if course.get('开始时间') and course.get('结束时间'):
                            course_text += f"\n时间：{course['开始时间']}~{course['结束时间']}"
                        if course.get('备注'):
                            course_text += f"\n备注：{course['备注']}"
                        course_texts.append(course_text)
                    
                    cell = ws.cell(row=row, column=col, value='\n'.join(course_texts))
                    cell.alignment = Alignment(wrap_text=True, vertical='center')
                    
                    # 获取课程颜色
                    first_course = courses[0]
                    color_rgb = get_course_color(first_course['课程名称'])
                    color_hex = f"{color_rgb[0]:02X}{color_rgb[1]:02X}{color_rgb[2]:02X}"
                    fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
                    cell.fill = fill
    
    # 设置列宽
    column_widths = [12, 15, 15, 15, 15, 15, 15, 15]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[chr(64 + i)].width = width
    
    # 设置行高
    for row in range(1, 16):
        if row in [3, 7, 11]:  # 上午/下午/晚自习行
            ws.row_dimensions[row].height = 25
        else:
            ws.row_dimensions[row].height = 40
    
    # 保存到输出流
    wb.save(output)

def generate_word(df, output, user_selected_colors, title):
    """生成Word文件"""
    # 创建文档
    doc = Document()
    
    # 设置页面方向为横向
    section = doc.sections[0]
    new_width, new_height = section.page_height, section.page_width
    section.orientation = 1  # 横向
    section.page_width = new_width
    section.page_height = new_height
    
    # 添加标题
    title_para = doc.add_paragraph()
    title_para.alignment = 1  # 居中
    title_run = title_para.add_run(title)
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    
    # 创建表格
    table = doc.add_table(rows=15, cols=8)
    table.style = 'Table Grid'
    
    # 设置标题行
    header_cells = table.rows[0].cells
    headers = ['节次/星期', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    for i, header in enumerate(headers):
        header_cells[i].text = header
        # 设置标题样式
        for paragraph in header_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
    
    # 合并标题行单元格
    header_cells[0].merge(header_cells[7])
    header_cells[0].text = title
    
    # 设置表头
    header_cells = table.rows[1].cells
    headers = ['节次/星期', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    for i, header in enumerate(headers):
        header_cells[i].text = header
        # 设置表头样式
        for paragraph in header_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
    
    # 设置上午/下午/晚自习标识
    time_periods = ['上午', '下午', '晚自习']
    period_rows = [2, 6, 10]  # 上午从第2行开始，下午从第6行开始，晚自习从第10行开始
    
    for i, (period, row_idx) in enumerate(zip(time_periods, period_rows)):
        period_row = table.rows[row_idx]
        period_cell = period_row.cells[0]
        period_cell.text = period
        period_cell.merge(period_row.cells[7])
        # 设置时段样式
        for paragraph in period_cell.paragraphs:
            paragraph.alignment = 1  # 居中
            for run in paragraph.runs:
                run.font.bold = True
    
    # 创建一个字典来快速查找课程
    course_dict = {}
    for _, course in df.iterrows():
        key = (course['星期'], course['节次'])
        if key not in course_dict:
            course_dict[key] = []
        course_dict[key].append(course)
    
    # 填充课程数据
    for period_rows_idx, start_row in enumerate([3, 7, 11]):  # 课程从第3、7、11行开始
        for row_offset in range(4):  # 每个时段4节课
            row = start_row + row_offset
            table_row = table.rows[row]
            
            # 设置节次
            period = row_offset + 1 + period_rows_idx * 4
            table_row.cells[0].text = f'第{period}节'
            
            # 填充每天的课程
            days = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
            for col, day in enumerate(days, 1):
                key = (day, period)
                if key in course_dict:
                    courses = course_dict[key]
                    # 如果同一时间有多门课程，显示所有课程
                    course_texts = []
                    for course in courses:
                        course_text = course['课程名称']
                        if course.get('教师') and course['教师'] != '未指定':
                            course_text += f"\n教师：{course['教师']}"
                        if course.get('地点') and course['地点'] != '未指定':
                            course_text += f"\n地点：{course['地点']}"
                        if course.get('开始时间') and course.get('结束时间'):
                            course_text += f"\n时间：{course['开始时间']}~{course['结束时间']}"
                        if course.get('备注'):
                            course_text += f"\n备注：{course['备注']}"
                        course_texts.append(course_text)
                    
                    table_row.cells[col].text = '\n'.join(course_texts)
    
    # 保存到输出流
    doc.save(output)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
