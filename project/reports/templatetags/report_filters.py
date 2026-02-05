# reports/templatetags/report_filters.py
from django import template
import re

register = template.Library()

@register.filter
def get_item(dictionary, key):
    """Получить значение из словаря по ключу"""
    if dictionary is None:
        return ''
    return dictionary.get(key, '')


@register.filter
def get_column_level(column_name):
    """
    Определить уровень столбца на основе имени
    Примеры:
    - '1' -> level 1
    - '1.2' -> level 2
    - '1.2.1' -> level 3
    - '12.2' -> level 2
    - 'mes' -> level 1
    """
    column_str = str(column_name)
    
    # Особые случаи
    if column_str == 'mes' or column_str == 'kod':
        return "1"
    
    # Подсчет точек для определения уровня
    dots = column_str.count('.')
    if dots == 0:
        return "1"
    elif dots == 1:
        return "2"
    else:
        return "3"


@register.filter
def get_mes_level(mes_value):

    if not mes_value:
        return 1
    
    mes_str = str(mes_value).strip()
    
    # Подсчет точек
    dots = mes_str.count('.')
    if dots == 0:
        return 1
    elif dots == 1:
        return 2
    else:
        return 3


@register.filter
def is_numeric(value):
    if value is None:
        return False
    
    value_str = str(value).strip()
    
    try:
        float(value_str)
        return True
    except ValueError:
        return False


@register.filter
def format_number(value):
    if not value:
        return value
    
    try:
        num = float(value)
        if num.is_integer():
            return "{:,.0f}".format(num).replace(',', ' ')
        else:
            return "{:,.2f}".format(num).replace(',', ' ')
    except (ValueError, TypeError):
        return value


@register.filter
def get_indent_style(level):
    level_int = int(level) if level else 1
    indent = (level_int - 1) * 1.5
    return f"padding-left: {indent}rem;"