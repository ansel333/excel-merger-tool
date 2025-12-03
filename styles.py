#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
样式表模块
"""

def get_stylesheet():
    """获取应用程序的样式表"""
    return """
        QMainWindow {
            background-color: #ffffff;
        }
        
        QWidget {
            background-color: #ffffff;
        }
        
        QLabel {
            color: #333333;
            font-size: 11px;
        }
        
        QPushButton {
            background-color: #5b9bd5;
            color: #ffffff;
            border: 1px solid #4a90c4;
            border-radius: 2px;
            padding: 5px 15px;
            font-size: 11px;
            min-height: 24px;
        }
        
        QPushButton:hover {
            background-color: #6ba3dc;
        }
        
        QPushButton:pressed {
            background-color: #4a90c4;
        }
        
        QPushButton:disabled {
            background-color: #cccccc;
            border: 1px solid #999999;
            color: #666666;
        }
        
        QListWidget {
            border: 1px solid #d0d0d0;
            background-color: #f9f9f9;
            border-radius: 2px;
        }
        
        QListWidget::item {
            padding: 2px 4px;
            margin: 0px;
            height: 18px;
            border-bottom: none;
        }
        
        QListWidget::item:selected {
            background-color: #5b9bd5;
            color: #ffffff;
        }
        
        QComboBox {
            border: 1px solid #d0d0d0;
            background-color: #ffffff;
            border-radius: 2px;
            padding: 4px;
            font-size: 11px;
        }
        
        QComboBox::drop-down {
            border-left: 1px solid #d0d0d0;
            width: 20px;
            background-color: #f5f5f5;
        }
        
        QComboBox QAbstractItemView {
            background-color: #ffffff;
            selection-background-color: #5b9bd5;
            selection-color: #ffffff;
        }
        
        QSpinBox {
            border: 1px solid #d0d0d0;
            background-color: #ffffff;
            border-radius: 2px;
            padding: 2px;
            font-size: 11px;
        }
        
        QSpinBox::up-button {
            subcontrol-origin: border;
            subcontrol-position: top right;
            width: 16px;
            border-left: 1px solid #d0d0d0;
            background-color: #f5f5f5;
        }
        
        QSpinBox::up-button:hover {
            background-color: #e0e0e0;
        }
        
        QSpinBox::down-button {
            subcontrol-origin: border;
            subcontrol-position: bottom right;
            width: 16px;
            border-left: 1px solid #d0d0d0;
            border-top: 1px solid #d0d0d0;
            background-color: #f5f5f5;
        }
        
        QSpinBox::down-button:hover {
            background-color: #e0e0e0;
        }
        
        QTextEdit {
            border: 1px solid #d0d0d0;
            background-color: #f0f4f8;
            border-radius: 2px;
            font-size: 11px;
            padding: 4px;
        }
        
        QScrollBar:vertical {
            background-color: #f5f5f5;
            width: 11px;
            border: none;
        }
        
        QScrollBar::handle:vertical {
            background-color: #b0b0b0;
            border-radius: 5px;
            min-height: 20px;
        }
        
        QScrollBar::handle:vertical:hover {
            background-color: #808080;
        }
        
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        
        QScrollBar:horizontal {
            background-color: #f5f5f5;
            height: 11px;
            border: none;
        }
        
        QScrollBar::handle:horizontal {
            background-color: #b0b0b0;
            border-radius: 5px;
            min-width: 20px;
        }
        
        QScrollBar::handle:horizontal:hover {
            background-color: #808080;
        }
        
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
            width: 0px;
        }
    """
