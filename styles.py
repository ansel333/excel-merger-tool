#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
样式表模块
"""

def get_stylesheet():
    """获取应用程序的样式表"""
    return """
        QMainWindow {
            background-color: #f0f4f8;
        }
        
        QWidget {
            background-color: #f0f4f8;
        }
        
        QLabel {
            color: #333333;
            font-size: 11px;
        }
        
        QLabel[heading="true"] {
            font-size: 10px;
            font-weight: bold;
            color: #4a90e2;
            background-color: transparent;
            padding: 4px 0px;
            border: none;
            border-radius: 0px;
        }
        
        QPushButton {
            background-color: #5b9bd5;
            color: #ffffff;
            border: 1px solid #4a90c4;
            border-radius: 2px;
            padding: 5px 12px;
            font-size: 10px;
            font-weight: 500;
            min-height: 26px;
        }
        
        QPushButton:hover {
            background-color: #6ba3dc;
            border: 1px solid #5b9bd5;
        }
        
        QPushButton:pressed {
            background-color: #4a90c4;
            border: 1px solid #3d7ab8;
        }
        
        QPushButton:disabled {
            background-color: #d0d0d0;
            color: #888888;
            border: 1px solid #b0b0b0;
        }
        
        QPushButton#btnPrimary {
            background-color: #5b9bd5;
            border: 1px solid #4a90c4;
        }
        
        QPushButton#btnPrimary:hover {
            background-color: #6ba3dc;
        }
        
        QPushButton#btnSecondary {
            background-color: #f0ad4e;
            border: 1px solid #e59c3d;
            color: #ffffff;
        }
        
        QPushButton#btnSecondary:hover {
            background-color: #f5bd5e;
            border: 1px solid #f0ad4e;
        }
        
        QPushButton#btnSecondary:pressed {
            background-color: #e59c3d;
            border: 1px solid #d48a2c;
        }
        
        QPushButton#btnSecondary:disabled {
            background-color: #d0d0d0;
            color: #888888;
            border: 1px solid #b0b0b0;
        }
        
        QPushButton#btnMerge {
            background-color: #5cb85c;
            border: 1px solid #4cae4c;
            min-height: 32px;
            font-size: 11px;
            font-weight: bold;
        }
        
        QPushButton#btnMerge:hover {
            background-color: #6cc96c;
            border: 1px solid #5cb85c;
        }
        
        QPushButton#btnMerge:pressed {
            background-color: #4cae4c;
            border: 1px solid #3d9e3d;
        }
        
        QComboBox {
            background-color: #ffffff;
            color: #333333;
            border: 1px solid #c0c0c0;
            border-radius: 2px;
            padding: 5px 8px;
            font-size: 10px;
            selection-background-color: #5b9bd5;
            selection-color: #ffffff;
        }
        
        QComboBox:focus {
            border: 1px solid #5b9bd5;
            background-color: #ffffff;
        }
        
        QComboBox#sheetCombo {
            background-color: #ffffff;
            color: #333333;
            border: 2px solid #4a90e2;
            border-radius: 2px;
            padding: 5px 8px;
            font-size: 10px;
            selection-background-color: #5b9bd5;
            selection-color: #ffffff;
        }
        
        QComboBox#sheetCombo:focus {
            border: 2px solid #3d7ab8;
            background-color: #ffffff;
        }
        
        QComboBox::drop-down {
            border: none;
            width: 18px;
        }
        
        QComboBox::down-arrow {
            image: none;
        }
        
        QComboBox QAbstractItemView {
            background-color: #ffffff;
            color: #333333;
            border: 1px solid #5b9bd5;
            selection-background-color: #5b9bd5;
            selection-color: #ffffff;
            outline: none;
        }
        
        QSpinBox {
            background-color: #ffffff;
            color: #333333;
            border: 1px solid #c0c0c0;
            border-radius: 2px;
            padding: 5px 8px;
            font-size: 10px;
        }
        
        QSpinBox:focus {
            border: 1px solid #5b9bd5;
            background-color: #ffffff;
        }
        
        QSpinBox#headerSpinBox {
            background-color: #ffffff;
            color: #333333;
            border: 2px solid #4a90e2;
            border-radius: 2px;
            padding: 5px 8px;
            font-size: 10px;
        }
        
        QSpinBox#headerSpinBox:focus {
            border: 2px solid #3d7ab8;
            background-color: #ffffff;
        }
        
        QListWidget {
            background-color: #ffffff;
            color: #333333;
            border: 1px solid #c0c0c0;
            border-radius: 2px;
            padding: 4px;
            font-size: 10px;
            outline: none;
        }
        
        QListWidget:focus {
            border: 1px solid #5b9bd5;
        }
        
        QListWidget#filesList {
            background-color: #ffffff;
            color: #333333;
            border: 2px solid #4a90e2;
            border-radius: 2px;
            padding: 4px;
            font-size: 10px;
            outline: none;
        }
        
        QListWidget#filesList:focus {
            border: 2px solid #3d7ab8;
        }
        
        QListWidget::item {
            padding: 2px 4px;
            border-radius: 1px;
            color: #333333;
            background-color: #ffffff;
            margin: 0px;
        }
        
        QListWidget::item:hover {
            background-color: #e8f0f8;
        }
        
        QListWidget::item:selected {
            background-color: #5b9bd5;
            color: #ffffff;
        }
        
        QTextEdit {
            background-color: #f5f9ff;
            color: #000000;
            border: 2px solid #4a90e2;
            border-radius: 3px;
            padding: 8px;
            font-size: 9px;
            font-family: 'Courier New', 'Consolas', monospace;
            outline: none;
            margin: 2px;
        }
        
        QTextEdit:focus {
            border: 2px solid #3d7ab8;
            background-color: #eef5ff;
        }
        
        QTextEdit#logBox {
            background-color: #ffffff;
            color: #000000;
            border: 2px solid #4a90e2;
            border-radius: 3px;
            padding: 8px;
            font-size: 9px;
            font-family: 'Courier New', 'Consolas', monospace;
            outline: none;
            margin: 2px;
        }
        
        QTextEdit#logBox:focus {
            border: 2px solid #3d7ab8;
            background-color: #f5f9ff;
        }
        
        QFrame {
            background-color: #f0f4f8;
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
