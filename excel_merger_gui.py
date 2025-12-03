#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel表格合并工具 - PyQt6 GUI版本
"""

import os
import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QFileDialog, QSpinBox,
    QLabel, QTextEdit, QMessageBox, QComboBox, QFrame, QAbstractSpinBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from styles import get_stylesheet
from merger import ExcelMerger


class ExcelMergerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_files = []
        self.sheet_names = []
        self.header_rows_count = 1
        self.selected_files = []
        self.init_ui()
        self.setStyleSheet(get_stylesheet())
    
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("Excel表格合并工具")
        self.setGeometry(100, 100, 850, 650)
        
        # 创建中夅窗口
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)
        
        # 标题
        title = QLabel("Excel表格合并工具")
        title_font = QFont()
        title_font.setPointSize(13)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setStyleSheet("color: #4a90e2;")
        main_layout.addWidget(title)
        
        # 分隔线
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.Shape.HLine)
        separator1.setStyleSheet("background-color: #d0d0d0;")
        separator1.setFixedHeight(1)
        main_layout.addWidget(separator1)
        
        # ========== 第一部分：文件选择 ==========
        file_section_layout = QHBoxLayout()
        file_section_layout.setSpacing(8)
        file_label = QLabel("步骤 1: 选择要合并的 Excel 文件")
        file_label.setProperty("heading", True)
        self.btn_select_files = QPushButton("📁 选择文件")
        self.btn_select_files.clicked.connect(self.select_files)
        self.btn_select_files.setObjectName("btnPrimary")
        file_section_layout.addWidget(file_label, 0)
        file_section_layout.addWidget(self.btn_select_files)
        file_section_layout.addStretch()
        main_layout.addLayout(file_section_layout)
        
        self.files_list = QListWidget()
        self.files_list.setObjectName("filesList")
        self.files_list.setMaximumHeight(90)
        self.files_list.setSpacing(0)
        main_layout.addWidget(self.files_list)
        
        # ========== 第二部分：读取Sheet ==========
        sheet_section_layout = QHBoxLayout()
        sheet_section_layout.setSpacing(8)
        sheet_label = QLabel("步骤 2: 读取 Sheet 列表")
        sheet_label.setProperty("heading", True)
        
        self.btn_read_sheets = QPushButton("📖 读取 Sheet")
        self.btn_read_sheets.clicked.connect(self.read_sheets)
        self.btn_read_sheets.setEnabled(False)
        self.btn_read_sheets.setObjectName("btnSecondary")
        sheet_section_layout.addWidget(sheet_label, 0)
        sheet_section_layout.addWidget(self.btn_read_sheets)
        sheet_section_layout.addStretch()
        main_layout.addLayout(sheet_section_layout)
        
        # ========== 第三部分：选择Sheet和表头行数 ==========
        config_section_layout = QVBoxLayout()
        config_section_layout.setSpacing(6)
        config_label = QLabel("步骤 3: 选择 Sheet 和设置表头行数")
        config_label.setProperty("heading", True)
        config_section_layout.addWidget(config_label)
        
        # Sheet选择和表头行数在同一行
        sheet_select_layout = QHBoxLayout()
        sheet_select_layout.setSpacing(8)
        sheet_select_layout.addWidget(QLabel("选择 Sheet:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setObjectName("sheetCombo")
        sheet_select_layout.addWidget(self.sheet_combo, 1)
        sheet_select_layout.addWidget(QLabel("表头行数:"), 0)
        self.header_spinbox = QSpinBox()
        self.header_spinbox.setObjectName("headerSpinBox")
        self.header_spinbox.setMinimum(1)
        self.header_spinbox.setValue(1)
        self.header_spinbox.setMaximumWidth(60)
        self.header_spinbox.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.UpDownArrows)
        sheet_select_layout.addWidget(self.header_spinbox, 0)
        config_section_layout.addLayout(sheet_select_layout)
        
        main_layout.addLayout(config_section_layout)
        
        # ========== 第四部分：合并 ==========
        merge_section_layout = QHBoxLayout()
        merge_section_layout.setSpacing(8)
        merge_label = QLabel("步骤 4: 执行合并")
        merge_label.setProperty("heading", True)
        
        self.btn_merge = QPushButton("✅ 开始合并")
        self.btn_merge.setObjectName("btnMerge")
        self.btn_merge.clicked.connect(self.merge_files)
        self.btn_merge.setEnabled(False)
        merge_section_layout.addWidget(merge_label, 0)
        merge_section_layout.addWidget(self.btn_merge)
        merge_section_layout.addStretch()
        
        main_layout.addLayout(merge_section_layout)
        
        # ========== 日志输出 ==========
        log_label = QLabel("📝 日志输出:")
        log_label.setProperty("heading", True)
        main_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setObjectName("logBox")
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(100)
        main_layout.addWidget(self.log_text, 1)  # 伸缩因子为 1，自动填充剩余空间
        
        central_widget.setLayout(main_layout)
    
    def select_files(self):
        """选择Excel文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择 Excel 文件",
            ".",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if files:
            self.selected_files = files
            
            # 显示选中的文件
            self.files_list.clear()
            for file in self.selected_files:
                self.files_list.addItem(os.path.basename(file))
            
            self.log(f"已选择 {len(self.selected_files)} 个文件")
            self.btn_read_sheets.setEnabled(True)
    
    def read_sheets(self):
        """读取选中文件的Sheet列表"""
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "请先选择文件！")
            return
        
        try:
            # 使用第一个文件获取sheet列表
            first_file = self.selected_files[0]
            self.sheet_names = ExcelMerger.get_sheet_names(first_file)
            
            # 更新Sheet选择框
            self.sheet_combo.clear()
            self.sheet_combo.addItems(self.sheet_names)
            
            self.log(f"成功读取 Sheet 列表，共 {len(self.sheet_names)} 个 Sheet")
            for i, sheet in enumerate(self.sheet_names, 1):
                self.log(f"  {i}. {sheet}")
            
            self.btn_merge.setEnabled(True)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取 Sheet 失败: {e}")
            self.log(f"错误: {e}")
    
    def merge_files(self):
        """执行合并"""
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "请先选择文件！")
            return
        
        if not self.sheet_names:
            QMessageBox.warning(self, "警告", "请先读取 Sheet 列表！")
            return
        
        target_sheet = self.sheet_combo.currentText()
        self.header_rows_count = self.header_spinbox.value()
        
        if not target_sheet:
            QMessageBox.warning(self, "警告", "请选择要合并的 Sheet！")
            return
        
        try:
            self.log(f"\n开始合并 Sheet: {target_sheet}")
            self.log(f"表头行数: {self.header_rows_count}")
            
            # 进行合并
            header_rows, data_rows, file_order = ExcelMerger.merge_sheets(
                self.selected_files, target_sheet, self.header_rows_count
            )
            
            if data_rows is None:
                QMessageBox.warning(self, "警告", "没有找到任何有效数据！")
                return
            
            self.log(f"合并完成！总共合并了 {len(data_rows)} 行数据")
            
            # 保存文件
            output_path = ExcelMerger.create_output_file(
                header_rows, data_rows, target_sheet, file_order, self.sheet_names, "."
            )
            
            if output_path:
                end_header_row = self.header_rows_count
                if end_header_row == 1:
                    self.log(f"  - 第1行: 表头（来自第一个文件）")
                else:
                    self.log(f"  - 第1-{end_header_row}行: 表头（来自第一个文件）")
                self.log(f"  - 第{end_header_row + 1}行起: 合并的数据行（共 {len(data_rows)} 行）")
            
            self.log(f"\n合并时使用的Excel文件顺序:")
            if file_order:
                for i, filename in enumerate(file_order, 1):
                    self.log(f"  {i}. {filename}")
            
            output_filename = os.path.basename(output_path) if output_path else "未知"
            QMessageBox.information(self, "成功", f"合并完成！\n文件已保存到: {output_filename}")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"合并失败: {e}")
            self.log(f"错误: {e}")
    
    def log(self, message):
        """输出日志"""
        self.log_text.append(message)
        # 自动滚动到最新位置
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )


def main():
    app = QApplication(sys.argv)
    window = ExcelMergerGUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
