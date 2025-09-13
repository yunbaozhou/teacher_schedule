def detect_conflicts(course_data):
    """Detect course conflicts: course arrangements for the same teacher at the same time"""
    conflicts = []
    # Group by teacher and time to check
    if "教师" in course_data.columns and "星期" in course_data.columns and "节次" in course_data.columns:
        grouped = course_data.groupby(["教师", "星期", "节次"])
        for key, group in grouped:
            if len(group) > 1:
                conflicts.append(f"冲突：教师{key[0]}在{key[1]}{key[2]}节有{len(group)}门课程")
    
    # Check course arrangements for the same class at the same time
    if "班级" in course_data.columns and "星期" in course_data.columns and "节次" in course_data.columns:
        grouped = course_data.groupby(["班级", "星期", "节次"])
        for key, group in grouped:
            if len(group) > 1:
                conflicts.append(f"冲突：班级{key[0]}在{key[1]}{key[2]}节有{len(group)}门课程")
    
    return conflicts