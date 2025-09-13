from flask import render_template

def index():
    """Main route, return frontend page"""
    return render_template('schedule.html')