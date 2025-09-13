from flask import jsonify

# Global variable to store course data
courses_data_store = []

def get_all_courses():
    """Get all courses"""
    return courses_data_store

def add_course(course_data):
    """Add a course"""
    try:
        courses_data_store.append(course_data)
        return {"success": True, "message": "课程添加成功"}
    except Exception as e:
        return {"success": False, "message": str(e)}

def clear_courses():
    """Clear all courses (useful for testing)"""
    courses_data_store.clear()