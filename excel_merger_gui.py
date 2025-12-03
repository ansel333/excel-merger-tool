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
    QLabel, QTextEdit, QMessageBox, QComboBox
)
from PyQt6.QtCore import Qt

from styles import get_stylesheet
from merger import ExcelMerger


class ExcelMergerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_files = []
        self.sheet_names = []
        self.header_rows_count = 1
        self.selected_files = []
        self.target_directory = None
        self.init_ui()
        self.setStyleSheet(get_stylesheet())
    
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("Excel表格合并工具")
        self.setGeometry(100, 100, 700, 550)
        
        # 创建中心窗口
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)
        
        # 步骤 1: 选择要合并的 Excel 文件
        step1_layout = QHBoxLayout()
        step1_layout.setSpacing(8)
        step1_label = QLabel("步骤 1: 选择要合并的 Excel 文件")
        self.btn_select_files = QPushButton("选择文件")
        self.btn_select_files.clicked.connect(self.select_files)
        step1_layout.addWidget(step1_label)
        step1_layout.addStretch()
        step1_layout.addWidget(self.btn_select_files)
        main_layout.addLayout(step1_layout)
        
        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(80)
        main_layout.addWidget(self.files_list)
        
        # 步骤 2: 读取 Sheet 列表
        step2_layout = QHBoxLayout()
        step2_layout.setSpacing(8)
        step2_label = QLabel("步骤 2: 读取 Sheet 列表")
        self.btn_read_sheets = QPushButton("读取 Sheet")
        self.btn_read_sheets.clicked.connect(self.read_sheets)
        self.btn_read_sheets.setEnabled(False)
        step2_layout.addWidget(step2_label)
        step2_layout.addStretch()
        step2_layout.addWidget(self.btn_read_sheets)
        main_layout.addLayout(step2_layout)
        
        # 步骤 3: 选择 Sheet 和设置表头行数
        step3_label = QLabel("步骤 3: 选择 Sheet 和设置表头行数")
        main_layout.addWidget(step3_label)
        
        sheet_select_layout = QHBoxLayout()
        sheet_select_layout.setSpacing(8)
        sheet_select_layout.addWidget(QLabel("选择 Sheet:"))
        self.sheet_combo = QComboBox()
        sheet_select_layout.addWidget(self.sheet_combo, 2)
        sheet_select_layout.addWidget(QLabel("表头行数:"))
        self.header_spinbox = QSpinBox()
        self.header_spinbox.setMinimum(1)
        self.header_spinbox.setValue(1)
        self.header_spinbox.setMinimumWidth(80)
        sheet_select_layout.addWidget(self.header_spinbox, 1)
        main_layout.addLayout(sheet_select_layout)
        
        # 步骤 4: 选择目标目录
        step4_label = QLabel("步骤 4: 选择目标目录")
        main_layout.addWidget(step4_label)
        
        target_dir_layout = QHBoxLayout()
        target_dir_layout.setSpacing(8)
        target_dir_layout.addWidget(QLabel("目标目录:"))
        self.target_dir_display = QLabel("未选择")
        self.target_dir_display.setStyleSheet("color: #999999;")
        target_dir_layout.addWidget(self.target_dir_display, 1)
        self.btn_select_dir = QPushButton("浏览...")
        self.btn_select_dir.clicked.connect(self.select_target_directory)
        self.btn_select_dir.setEnabled(False)
        target_dir_layout.addWidget(self.btn_select_dir)
        main_layout.addLayout(target_dir_layout)
        
        # 步骤 5: 执行合并
        step5_layout = QHBoxLayout()
        step5_layout.setSpacing(8)
        step5_label = QLabel("步骤 5: 执行合并")
        self.btn_merge = QPushButton("开始合并")
        self.btn_merge.clicked.connect(self.merge_files)
        self.btn_merge.setEnabled(False)
        step5_layout.addWidget(step5_label)
        step5_layout.addStretch()
        step5_layout.addWidget(self.btn_merge)
        main_layout.addLayout(step5_layout)
        
        # 日志输出
        log_label = QLabel("日志:")
        main_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(150)
        main_layout.addWidget(self.log_text, 1)
        
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
            
            # 设置默认目标目录为第一个文件所在的目录
            default_dir = os.path.dirname(os.path.abspath(first_file))
            self.target_directory = default_dir
            self.target_dir_display.setText(default_dir)
            self.target_dir_display.setStyleSheet("color: #333333;")
            self.log(f"默认目标目录: {default_dir}")
            
            # 启用目录选择按钮
            self.btn_select_dir.setEnabled(True)
            self.update_merge_button_state()
            
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
        
        if not self.target_directory:
            QMessageBox.warning(self, "警告", "请先选择目标目录！")
            return
        
        target_sheet = self.sheet_combo.currentText()
        self.header_rows_count = self.header_spinbox.value()
        
        if not target_sheet:
            QMessageBox.warning(self, "警告", "请选择要合并的 Sheet！")
            return
        
        try:
            self.log(f"\n开始合并 Sheet: {target_sheet}")
            self.log(f"表头行数: {self.header_rows_count}")
            self.log(f"目标目录: {self.target_directory}")
            
            # 进行合并
            header_rows, data_rows, file_order = ExcelMerger.merge_sheets(
                self.selected_files, target_sheet, self.header_rows_count
            )
            
            if data_rows is None:
                QMessageBox.warning(self, "警告", "没有找到任何有效数据！")
                return
            
            self.log(f"合并完成！总共合并了 {len(data_rows)} 行数据")
            
            # 保存文件到目标目录
            output_path = ExcelMerger.create_output_file(
                header_rows, data_rows, target_sheet, file_order, self.sheet_names, self.target_directory
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
    
    def select_target_directory(self):
        """选择目标目录"""
        directory = QFileDialog.getExistingDirectory(
            self,
            "选择目标目录",
            self.target_directory or "."
        )
        
        if directory:
            self.target_directory = directory
            self.target_dir_display.setText(directory)
            self.target_dir_display.setStyleSheet("color: #333333;")
            self.log(f"已选择目标目录: {directory}")
            self.update_merge_button_state()
    
    def update_merge_button_state(self):
        """更新合并按钮的状态"""
        can_merge = (
            self.selected_files and 
            self.sheet_names and 
            self.target_directory is not None
        )
        self.btn_merge.setEnabled(can_merge)
    
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
