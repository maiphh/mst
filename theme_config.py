# Theme Configuration for Talentnet TaxChecker
# Adjust these colors to customize the app appearance

# Brand Colors (#B74626 base)
BRAND = {
    "50": "#fdf4f2",
    "100": "#fbe6e2",
    "200": "#f6d0c8",
    "300": "#efa89a",
    "400": "#e57a64",
    "500": "#d65231",
    "600": "#b74626",  # Primary brand color
    "700": "#98361e",
    "800": "#7e301e",
    "900": "#682b1d",
    "950": "#38130b",
}

# Neutral Colors (for backgrounds and general UI)
NEUTRAL = {
    "white": "#ffffff",
    "grey_50": "#fafafa",
    "grey_100": "#f5f5f5",
    "grey_200": "#eeeeee",
    "grey_300": "#e0e0e0",
    "grey_400": "#bdbdbd",
    "grey_500": "#9e9e9e",
    "grey_600": "#757575",
    "grey_700": "#616161",
    "grey_800": "#424242",
    "grey_900": "#212121",
}

# Neutral Colors (refined for UX)
NEUTRAL = {
    "white": "#ffffff",
    "surface": "#f8f9fa",      # Slightly off-white for background
    "surface_alt": "#f1f3f5",  # For specialized areas
    
    # Text colors
    "text_main": "#343a40",    # Soft black/Dark grey
    "text_para": "#495057",    # Grey text
    "text_mute": "#868e96",    # Muted text
    
    # Greys for UI elements
    "grey_border": "#dee2e6",
    "grey_light": "#e9ecef",
    "grey_med": "#ced4da",
}

# Theme mappings - Neutral backgrounds, Brand for attention/buttons
THEME = {
    # Backgrounds (neutral white-grey)
    "background": NEUTRAL["surface"],         # Changed from pure white
    "background_alt": NEUTRAL["surface_alt"],
    
    # Text colors (softer black)
    "text_primary": NEUTRAL["text_main"],
    "text_secondary": NEUTRAL["text_para"],
    "text_muted": NEUTRAL["text_mute"],
    
    # Button colors - Primary (Brand)
    "button_primary": BRAND["600"],
    "button_hover": BRAND["700"],
    "button_pressed": BRAND["800"],
    "button_text": NEUTRAL["white"],
    
    # Button colors - Secondary (Grey/Neutral)
    "button_sec_bg": NEUTRAL["grey_light"],
    "button_sec_text": NEUTRAL["text_main"],
    "button_sec_hover": NEUTRAL["grey_med"],
    "button_sec_pressed": NEUTRAL["grey_border"],
    
    # Disabled state
    "button_disabled_bg": NEUTRAL["grey_light"],
    "button_disabled_text": NEUTRAL["text_mute"],
    
    # Border colors (neutral)
    "border": NEUTRAL["grey_border"],
    "border_focus": BRAND["600"],
    
    # Progress bar (brand for attention)
    "progress_bg": NEUTRAL["grey_light"],
    "progress_fill": BRAND["600"],
    
    # Input fields (neutral)
    "input_bg": NEUTRAL["white"],
    "input_border": NEUTRAL["grey_border"],
}


def get_app_stylesheet():
    """Generate the complete app stylesheet using theme colors."""
    return f"""
        QMainWindow {{
            background-color: {THEME["background"]};
        }}
        QWidget {{
            background-color: {THEME["background"]};
            font-family: Arial;
        }}
        QLabel {{
            color: {THEME["text_primary"]};
        }}
        QPushButton {{
            background-color: {THEME["button_primary"]};
            color: {THEME["button_text"]};
            border: none;
            padding: 0 16px;
            border-radius: 6px;
            font-weight: bold;
            min-height: 40px;
        }}
        QPushButton:hover {{
            background-color: {THEME["button_hover"]};
        }}
        QPushButton:pressed {{
            background-color: {THEME["button_pressed"]};
        }}
        QPushButton:disabled {{
            background-color: {THEME["button_disabled_bg"]};
            color: {THEME["button_disabled_text"]};
        }}
        
        QProgressBar {{
            border: 2px solid {THEME["border"]};
            border-radius: 10px;
            background-color: {THEME["progress_bg"]};
            text-align: center;
        }}
        QProgressBar::chunk {{
            background-color: {THEME["progress_fill"]};
            border-radius: 8px;
        }}
        QGroupBox {{
            font-weight: bold;
            border: 1px solid {THEME["border"]};
            border-radius: 10px;
            margin-top: 10px;
            padding-top: 10px;
            background-color: {THEME["background_alt"]};
        }}
        QGroupBox::title {{
            color: {THEME["button_primary"]};
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px;
        }}
        QTextEdit {{
            border: 1px solid {THEME["border"]};
            border-radius: 10px;
            background-color: {THEME["input_bg"]};
            color: {THEME["text_primary"]};
        }}
        QSpinBox {{
            border: 1px solid {THEME["border"]};
            border-radius: 6px;
            padding: 0 10px;
            background: {THEME["input_bg"]};
            color: {THEME["text_primary"]};
            min-height: 40px;
        }}
        QSpinBox::up-button, QSpinBox::down-button {{
            width: 0px;
            height: 0px;
            border: none;
            image: none;
        }}
        
        QCheckBox {{
            spacing: 8px;
            color: {THEME["text_primary"]};
        }}
        QCheckBox::indicator {{
            width: 18px;
            height: 18px;
            border: 1px solid {THEME["border"]};
            border-radius: 4px;
            background-color: {THEME["input_bg"]};
        }}
        QCheckBox::indicator:hover {{
            border-color: {THEME["border_focus"]};
        }}
        QCheckBox::indicator:checked {{
            background-color: {THEME["button_primary"]};
            border: 1px solid {THEME["button_primary"]};
            
        }}
    """


def get_splash_progressbar_style():
    """Get the progress bar style for splash screen."""
    return f"""
        QProgressBar {{
            border: 2px solid {THEME["border"]};
            border-radius: 10px;
            background-color: {THEME["progress_bg"]};
            color: {THEME["text_primary"]};
        }}
        QProgressBar::chunk {{
            background-color: {THEME["progress_fill"]};
            border-radius: 8px;
        }}
    """


def get_input_label_style():
    """Get the style for input path labels."""
    return f"border: 1px solid {THEME['border']}; padding: 0 12px; background: {THEME['input_bg']}; color: {THEME['text_primary']}; border-radius: 6px; min-height: 40px;"


def get_secondary_button_style():
    """Get stylesheet for secondary action buttons (grey)."""
    return f"""
        QPushButton {{
            background-color: {THEME["button_sec_bg"]};
            color: {THEME["button_sec_text"]};
            border: 1px solid {THEME["border"]};
            border-radius: 6px;
            font-weight: 600;
            min-height: 40px;
        }}
        QPushButton:hover {{
            background-color: {THEME["button_sec_hover"]};
            border: 1px solid {NEUTRAL["grey_med"]};
        }}
        QPushButton:pressed {{
            background-color: {THEME["button_sec_pressed"]};
        }}
        QPushButton:disabled {{
            background-color: {THEME["button_disabled_bg"]};
            color: {THEME["button_disabled_text"]};
            border: none;
        }}
    """
