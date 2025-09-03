import pandas as pd
from docx import Document
from docx.shared import Inches, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
import json
import random
from flask import Flask, request, jsonify, send_file, render_template
import os

app = Flask(__name__, static_folder='.', static_url_path='')

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
]

# 用于存储课程数据的全局变量
courses_data_store = []

# 根路由，返回前端页面
@app.route('/')
def index():
    return send_file('schedule.html')

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
        df = pd.DataFrame(courses)
        
        # 生成Excel文件
        filename = "课程表.xlsx"
        filepath = generate_excel(df, filename)
        
        # 返回文件
        return send_file(filepath, as_attachment=True)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400

@app.route('/api/export/word', methods=['POST'])
def export_word():
    """导出为Word"""
    try:
        data = request.json
        courses = data.get('courses', [])
        user_selected_colors = data.get('userSelectedColors', {})
        df = pd.DataFrame(courses)
        
        # 生成Word文件
        filename = "课程表.docx"
        filepath = generate_word(df, filename, user_selected_colors)
        
        # 返回文件
        return send_file(filepath, as_attachment=True)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400

def get_course_color(course_name, user_selected_colors=None):
    """
    获取课程颜色，优先级：
    1. 用户选择的颜色
    2. 预定义的颜色映射
    3. 自动分配默认颜色
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
    
    # 如果都没有，则从默认颜色中自动分配
    # 为了确保相同课程名获得相同颜色，我们基于课程名生成一个索引
    course_hash = hash(course_name) % len(DEFAULT_COLORS)
    return DEFAULT_COLORS[course_hash]

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
    if "星期" in course_data.columns and "节次" in course_data.columns:
        grouped_class = course_data.groupby(["星期", "节次"])
        for key, group in grouped_class:
            if len(group) > 1:
                # 这里假设所有数据都是同一个班级的
                conflicts.append(f"冲突：班级在{key[0]}{key[1]}节有{len(group)}门课程")
    
    return conflicts

def generate_excel(course_data, output_file="课程表.xlsx"):
    """生成Excel格式课程表"""
    # 创建透视表，包含更多信息
    try:
        # 确保必要的列存在
        required_columns = ["节次", "星期", "课程名称"]
        if all(col in course_data.columns for col in required_columns):
            pivot_table = course_data.pivot_table(
                index="节次", 
                columns="星期", 
                values=["课程名称", "教师", "地点"], 
                aggfunc=lambda x: '\n'.join(str(v) for v in x)
            )
            pivot_table.to_excel(output_file)
        else:
            # 如果缺少必要列，使用简单版本
            pivot = course_data.pivot(index="节次", columns="星期", values="课程名称")
            pivot.to_excel(output_file)
    except Exception as e:
        # 如果透视表创建失败，使用简单版本
        try:
            pivot = course_data.pivot(index="节次", columns="星期", values="课程名称")
            pivot.to_excel(output_file)
        except Exception as e2:
            # 最后的备选方案
            course_data.to_excel(output_file, index=False)
    
    return output_file

def set_cell_background_color(cell, color):
    """设置单元格背景颜色"""
    # 将RGB颜色值转换为十六进制
    hex_color = '{:02x}{:02x}{:02x}'.format(*color)
    
    # 创建XML元素来设置背景色
    from docx.oxml.shared import OxmlElement, qn
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    tc_pr.append(shd)

def generate_word(course_data, output_file="课程表.docx", user_selected_colors=None):
    """生成Word格式课程表"""
    doc = Document()
    doc.add_heading("课程表", level=1)
    
    # 确保必要的列存在
    required_columns = ["节次", "星期", "课程名称"]
    if not all(col in course_data.columns for col in required_columns):
        doc.add_paragraph("数据格式不正确，缺少必要列")
        doc.save(output_file)
        return output_file
    
    # 创建表格：行数=节次+1，列数=星期+1
    try:
        max_period = int(course_data["节次"].max())
    except:
        max_period = 8  # 默认最大节次
    
    # 确保星期列按正确顺序排列
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    available_weekdays = [day for day in weekdays if day in course_data["星期"].unique()]
    
    table = doc.add_table(rows=max_period+1, cols=len(available_weekdays)+1)
    table.style = 'Table Grid'
    
    # 设置表头
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "节次"
    for i, day in enumerate(available_weekdays):
        hdr_cells[i+1].text = day
    
    # 填充内容
    for period in range(1, max_period+1):
        row_cells = table.rows[period].cells
        row_cells[0].text = str(period)
        for i, day in enumerate(available_weekdays):
            courses = course_data[(course_data["节次"] == period) & (course_data["星期"] == day)]
            if not courses.empty:
                # 合并课程信息
                cell_text = ""
                for _, course in courses.iterrows():
                    course_name = course.get("课程名称", "未知课程")
                    teacher = course.get("教师", "未知教师")
                    location = course.get("地点", "未知地点")
                    cell_text += f"{course_name}\n{teacher}\n{location}\n\n"
                
                # 删除末尾多余的换行
                cell_text = cell_text.rstrip()
                row_cells[i+1].text = cell_text
                
                # 为第一个课程设置背景色
                first_course = courses.iloc[0]
                course_name = first_course.get("课程名称", "未知课程")
                course_color = get_course_color(course_name, user_selected_colors)
                set_cell_background_color(row_cells[i+1], course_color)
    
    doc.save(output_file)
    return output_file

def load_course_data_from_excel(file_path):
    """
    从Excel文件加载课程数据
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return None

def load_course_data_from_json(file_path):
    """
    从JSON文件加载课程数据
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        df = pd.DataFrame(data)
        return df
    except Exception as e:
        print(f"读取JSON文件失败: {e}")
        return None

def main():
    # 模拟输入数据（实际应用中从前端或导入文件获取）
    data = {
        "课程名称": ["语文", "数学", "英语", "体育", "音乐", "美术"],
        "教师": ["张三", "李四", "王五", "赵六", "钱七", "孙八"],
        "星期": ["周一", "周二", "周三", "周四", "周五", "周一"],
        "节次": [1, 2, 3, 4, 5, 1],
        "地点": ["101教室", "102教室", "103教室", "操场", "音乐室", "美术室"]
    }
    course_data = pd.DataFrame(data)
    
    # 确保节次是数字类型
    course_data["节次"] = pd.to_numeric(course_data["节次"], errors='coerce')
    
    # 模拟用户选择的颜色（实际应用中从前端获取）
    user_selected_colors = {
        # "语文": (255, 0, 0),  # 红色，示例
    }
    
    # 检测冲突
    conflicts = detect_conflicts(course_data)
    if conflicts:
        print("检测到课程冲突：")
        for conflict in conflicts:
            print(conflict)
    else:
        print("未检测到课程冲突")
    
    # 生成Excel和Word
    generate_excel(course_data)
    generate_word(course_data, user_selected_colors=user_selected_colors)
    
    print("数据处理完成")

if __name__ == "__main__":
    # 运行Flask应用
    app.run(debug=True, port=5000)