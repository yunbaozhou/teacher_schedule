import pandas as pd
from flask import Flask
from config import FLASK_CONFIG
from models import courses_data_store
from routes.main_routes import index
from routes.api_routes import get_courses, add_course_endpoint, check_conflicts, export_excel, export_word

app = Flask(__name__, 
            static_folder=FLASK_CONFIG['static_folder'], 
            template_folder=FLASK_CONFIG['template_folder'])

# Main route
@app.route('/')
def index_route():
    return index()

# API routes
@app.route('/api/courses', methods=['GET'])
def get_courses_route():
    return get_courses()

@app.route('/api/courses', methods=['POST'])
def add_course_route():
    return add_course_endpoint()

@app.route('/api/courses/conflicts', methods=['POST'])
def check_conflicts_route():
    return check_conflicts()

@app.route('/api/export/excel', methods=['POST'])
def export_excel_route():
    return export_excel()

@app.route('/api/export/word', methods=['POST'])
def export_word_route():
    return export_word()

# Start Flask application
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
