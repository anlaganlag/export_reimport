#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Module for unit conversion utilities.
"""

def translate_unit(unit):
    """
    Translate unit abbreviations to standardized unit names.
    This is extracted from the original process_shipping_list.py.

    Args:
        unit (str): Input unit string

    Returns:
        str: Standardized unit string
    """
    if unit is None:
        return "PCS"
        
    unit = str(unit).strip().upper()
    
    # Common English units
    if unit in ["PCS", "PC", "PIECE", "PIECES", "EA", "EACH"]:
        return "PCS"
    elif unit in ["SET", "SETS"]:
        return "SET"
    elif unit in ["M", "MTR", "METER", "METERS"]:
        return "M"
    elif unit in ["KG", "KGS", "KILOGRAM", "KILOGRAMS"]:
        return "KG"
    elif unit in ["G", "GRAM", "GRAMS"]:
        return "G"
    elif unit in ["T", "TON", "TONS", "TONNE", "TONNES"]:
        return "T"
    elif unit in ["L", "LTR", "LITER", "LITERS", "LITRE", "LITRES"]:
        return "L"
    elif unit in ["ML", "MILLILITER", "MILLILITERS", "MILLILITRE", "MILLILITRES"]:
        return "ML"
    
    # Common Chinese units
    elif unit in ["个", "件", "只", "台", "支"]:
        return "PCS"
    elif unit in ["套"]:
        return "SET"
    elif unit in ["米"]:
        return "M"
    elif unit in ["千克", "公斤"]:
        return "KG"
    elif unit in ["克"]:
        return "G"
    elif unit in ["吨"]:
        return "T"
    elif unit in ["升", "立升"]:
        return "L"
    elif unit in ["毫升"]:
        return "ML"
    
    # If no match found, return the original unit
    return unit 