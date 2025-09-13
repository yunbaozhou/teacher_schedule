import pandas as pd
from io import BytesIO
from config import EXPORT_CONFIG
from services.color_service import get_course_color

class BaseExportService:
    """Base export service class"""
    
    def __init__(self):
        pass
    
    def prepare_data(self, data):
        """Prepare data for export"""
        if not data:
            raise ValueError("请求数据为空")
            
        courses = data.get('courses', [])
        title = data.get('title', EXPORT_CONFIG['default_title'])
        
        if not courses:
            raise ValueError("没有课程数据可供导出")
        
        df = pd.DataFrame(courses)
        return df, title, data