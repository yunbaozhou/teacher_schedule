import pandas as pd
from flask import jsonify, request, send_file
from io import StringIO
import sys
from models import get_all_courses, add_course
from services.conflict_service import detect_conflicts
# 从新的服务文件导入
from services.excel_export_service import ExcelExportService
from services.word_export_service import WordExportService
from services.statistics_service import statistics_service

# 创建服务实例
excel_export_service = ExcelExportService()
word_export_service = WordExportService()

def get_courses():
    """Get all courses"""
    return jsonify(get_all_courses())

def add_course_endpoint():
    """Add course"""
    try:
        course_data = request.json
        result = add_course(course_data)
        if result["success"]:
            return jsonify(result)
        else:
            return jsonify(result), 400
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400

def check_conflicts():
    """Check course conflicts"""
    try:
        # Save original stdout
        old_stdout = sys.stdout
        sys.stdout = StringIO()
        
        data = request.json
        courses = data.get('courses', [])
        df = pd.DataFrame(courses)
        conflicts = detect_conflicts(df)
        
        # Restore original stdout
        sys.stdout = old_stdout
        
        return jsonify({"conflicts": conflicts})
    except Exception as e:
        return jsonify({"conflicts": [], "error": str(e)}), 400

def export_excel():
    """Export to Excel"""
    try:
        # Track statistics
        statistics_service.track_export("excel")
        
        data = request.json
        output, filename = excel_export_service.create_excel_export(data)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except ValueError as e:
        return jsonify({"success": False, "message": str(e)}), 400
    except Exception as e:
        # Record detailed error information
        import traceback
        error_info = traceback.format_exc()
        print(f"Error exporting Excel: {str(e)}")
        print(f"Detailed error information:\n{error_info}")
        return jsonify({"success": False, "message": f"导出失败: {str(e)}"}), 500

def export_word():
    """Export to Word"""
    try:
        # Track statistics
        statistics_service.track_export("word")
        
        data = request.json
        output, filename = word_export_service.create_word_export(data)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except ValueError as e:
        return jsonify({"success": False, "message": str(e)}), 400
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400

def export_image():
    """Track image export"""
    # Track statistics
    statistics_service.track_export("image")
    return jsonify({"success": True})

def print_schedule():
    """Track print action"""
    # Track statistics
    statistics_service.track_export("print")
    return jsonify({"success": True})

def import_schedule():
    """Track import action"""
    # Track statistics
    statistics_service.track_import()
    return jsonify({"success": True})

def get_statistics():
    """Get usage statistics"""
    stats = statistics_service.get_stats()
    # Ensure all required fields exist
    if "export_stats" not in stats:
        stats["export_stats"] = {"excel": 0, "word": 0, "image": 0, "print": 0}
    if "import_stats" not in stats:
        stats["import_stats"] = {"total": 0}
        
    return jsonify(stats)