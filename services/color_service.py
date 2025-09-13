from config import COLOR_MAP, DEFAULT_COLORS

def get_course_color(course_name, user_selected_colors=None):
    """
    Get course color, priority:
    1. User selected colors
    2. Predefined color mapping
    3. Automatically assign default colors (based on strict course name calculation)
    """
    # If the user has selected a color, use the user selected color first
    if user_selected_colors and course_name in user_selected_colors:
        color = user_selected_colors[course_name]
        # If the color is in string format, convert it to a tuple
        if isinstance(color, str):
            return tuple(map(int, color.split(',')))
        return color
    
    # If a predefined color exists, use the predefined color
    if course_name in COLOR_MAP:
        return COLOR_MAP[course_name]
    
    # Generate a stable hash value based on the course name to ensure that 
    # the same name course always gets the same color
    # Use a larger color pool to reduce the probability of different name courses getting the same color
    hash_value = hash(course_name) % (len(DEFAULT_COLORS) * 100)  # Expand color space
    color_index = hash_value % len(DEFAULT_COLORS)
    return DEFAULT_COLORS[color_index]