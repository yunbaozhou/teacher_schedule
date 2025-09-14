from flask import render_template
from models import courses_data_store

def index():
    """Main page"""
    return render_template('schedule.html', courses=courses_data_store)