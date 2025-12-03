#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel表格合并工具
功能：扫描目录下的所有Excel文件，让用户选择要合并的sheet，将多个文件的相同sheet合并为一个新文件
"""

import os
import glob
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import re

class ExcelMerger:
    def __init__(self, directory="."):
        """
        初始化Excel合并器
        
        Args:
            directory (str): 要扫描的目录路径，默认为当前目录
        """
        self.directory = directory
        self.excel_files = []
        self.sheet_names = []
        self.header_rows_count = 1  # 表头行数，默认为1行
        
    def scan_excel_files(self):
        """扫描目录下的所有Excel文件"""
        print(f"正在扫描目录: {os.path.abspath(self.directory)}")
        
        # 查找所有xlsx和xls文件
        xlsx_files = glob.glob(os.path.join(self.directory, "*.xlsx"))
        xls_files = glob.glob(os.path.join(self.directory, "*.xls"))
        
        all_files = xlsx_files + xls_files
        
        # 过滤掉包含"合并结果"的文件
        all_files = [f for f in all_files if "合并结果" not in os.path.basename(f)]
        
        if not all_files:
            print("未找到任何Excel文件！")
            return False
            
        # 按文件修改时间排序
        self.excel_files = sorted(all_files, key=lambda x: os.path.getmtime(x))
        
        print(f"找到 {len(self.excel_files)} 个Excel文件:")
        for i, file in enumerate(self.excel_files, 1):
            mtime = datetime.fromtimestamp(os.path.getmtime(file))
            print(f"  {i}. {os.path.basename(file)} (修改时间: {mtime.strftime('%Y-%m-%d %H:%M:%S')})")
        
        return True
    
    def get_sheet_names(self):
        """获取第一个Excel文件中的所有sheet名称"""
        if not self.excel_files:
            return False
            
        try:
            # 使用第一个文件获取sheet列表
            first_file = self.excel_files[0]
            xl_file = pd.ExcelFile(first_file)
            self.sheet_names = xl_file.sheet_names
            xl_file.close()
            
            print(f"\n在文件 '{os.path.basename(first_file)}' 中找到以下sheet:")
            for i, sheet in enumerate(self.sheet_names, 1):
                print(f"  {i}. {sheet}")
            
            return True
        except Exception as e:
            print(f"读取Excel文件时出错: {e}")
            return False
    
    def select_sheet(self):
        """让用户选择要合并的sheet"""
        while True:
            try:
                choice = input(f"\n请选择要合并的sheet (1-{len(self.sheet_names)}): ").strip()
                
                if choice.isdigit():
                    index = int(choice) - 1
                    if 0 <= index < len(self.sheet_names):
                        selected_sheet = self.sheet_names[index]
                        print(f"您选择了sheet: {selected_sheet}")
                        return selected_sheet
                
                print("请输入有效的数字！")
                
            except KeyboardInterrupt:
                print("\n操作已取消")
                return None
    
    def select_header_rows(self):
        """让用户输入表头行数"""
        while True:
            try:
                choice = input(f"\n请输入表头所占的行数 (默认为1，表示只有表头行): ").strip()
                
                if not choice:
                    self.header_rows_count = 1
                    print(f"使用默认值: 表头 {self.header_rows_count} 行")
                    return True
                
                if choice.isdigit():
                    num = int(choice)
                    if num >= 1:
                        self.header_rows_count = num
                        print(f"已设置: 表头 {self.header_rows_count} 行")
                        return True
                    else:
                        print("请输入大于等于1的数字！")
                else:
                    print("请输入有效的数字！")
                    
            except KeyboardInterrupt:
                print("\n操作已取消")
                return False
    def merge_sheets(self, target_sheet_name):
        """
        合并所有Excel文件中的指定sheet
        跳过前两行（标题和表头），只合并数据行
        
        Args:
            target_sheet_name (str): 要合并的sheet名称
        """
        print(f"\n开始合并sheet: {target_sheet_name}")
        
        all_data = []
        file_order = []
        header_rows = None  # 保存标题和表头
        header_found = False  # 标记是否已找到表头
        
        for file_path in self.excel_files:
            try:
                print(f"正在处理文件: {os.path.basename(file_path)}")
                
                # 读取指定sheet，不跳过任何行，先获取完整数据
                df_full = pd.read_excel(file_path, sheet_name=target_sheet_name, header=None)
                
                if df_full.empty:
                    print(f"  警告: sheet '{target_sheet_name}' 在文件中为空")
                    continue
                
                # 保存第一个有效文件的表头行
                if not header_found:
                    if len(df_full) >= self.header_rows_count:
                        header_rows = df_full.iloc[:self.header_rows_count].copy()
                        header_found = True
                        print(f"  保存表头行（前 {self.header_rows_count} 行）")
                    else:
                        print(f"  警告: 文件行数不足 {self.header_rows_count} 行，跳过此文件以寻找有效表头")
                        continue
                
                # 提取从表头之后开始的数据
                data_start_row = self.header_rows_count
                if len(df_full) > data_start_row:
                    data_rows = df_full.iloc[data_start_row:].copy()
                    data_rows.reset_index(drop=True, inplace=True)
                    
                    if not data_rows.empty:
                        all_data.append(data_rows)
                        file_order.append(os.path.basename(file_path))
                        print(f"  提取了 {len(data_rows)} 行数据")
                    else:
                        print(f"  警告: 表头后没有数据")
                elif len(df_full) == self.header_rows_count:
                    print(f"  信息: 文件只有表头，没有数据行")
                else:
                    print(f"  警告: 文件行数不足2行，无法提取数据行")
                    
            except Exception as e:
                print(f"  错误: 无法读取文件 {os.path.basename(file_path)} 的sheet '{target_sheet_name}': {e}")
                continue
        
        if not all_data:
            print("没有找到任何有效数据！")
            return None, None, None
        
        if not header_found:
            print("警告: 未找到有效的表头行（所有文件行数都不足2行）")
            return None, None, None
        
        # 合并所有数据行
        merged_data = pd.concat(all_data, ignore_index=True)
        
        # 如果第一列是编号列，重新编号
        if not merged_data.empty and header_rows is not None:
            # 检查第一列是否可能是编号列
            first_col_data = merged_data.iloc[:, 0]
            # 检查是否为数值类型或包含数字的字符串
            is_numeric_col = (
                first_col_data.dtype in ['int64', 'float64'] or
                first_col_data.astype(str).str.match(r'^\d+$', na=False).any()
            )
            
            if is_numeric_col:
                merged_data.iloc[:, 0] = range(1, len(merged_data) + 1)
                print(f"已重新编号第一列，编号范围: 1-{len(merged_data)}")
        
        print(f"合并完成！总共合并了 {len(merged_data)} 行数据")
        
        return header_rows, merged_data, file_order
    
    def create_output_file(self, header_rows, merged_data, target_sheet_name, file_order):
        """
        创建输出Excel文件
        
        Args:
            header_rows (DataFrame): 标题和表头行（前两行）
            merged_data (DataFrame): 合并后的数据行
            target_sheet_name (str): 目标sheet名称
            file_order (list): 文件处理顺序
        """
        # 生成输出文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"合并结果_{target_sheet_name}_{timestamp}.xlsx"
        output_path = os.path.join(self.directory, output_filename)
        
        try:
            # 创建新的工作簿
            wb = openpyxl.Workbook()
            
            # 删除默认sheet
            wb.remove(wb.active)
            
            # 为每个原有sheet创建空sheet
            for sheet_name in self.sheet_names:
                ws = wb.create_sheet(title=sheet_name)
                
                if sheet_name == target_sheet_name:
                    # 写入标题和表头行
                    if header_rows is not None:
                        for _, row in header_rows.iterrows():
                            ws.append(row.tolist())
                    
                    # 写入合并的数据行
                    for _, row in merged_data.iterrows():
                        ws.append(row.tolist())
                    
                    # 调整列宽
                    for column in ws.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        ws.column_dimensions[column_letter].width = adjusted_width
            
            # 保存文件
            wb.save(output_path)
            wb.close()
            
            print(f"\n合并结果已保存到: {output_filename}")
            print(f"文件结构:")
            end_header_row = self.header_rows_count
            if end_header_row == 1:
                print(f"  - 第1行: 表头（来自第一个文件）")
            else:
                print(f"  - 第1-{end_header_row}行: 表头（来自第一个文件）")
            print(f"  - 第{end_header_row + 1}行起: 合并的数据行（共 {len(merged_data)} 行）")
            
            # 输出文件处理顺序
            print(f"\n合并时使用的Excel文件顺序:")
            for i, filename in enumerate(file_order, 1):
                print(f"  {i}. {filename}")
            
            return output_path
            
        except Exception as e:
            print(f"保存文件时出错: {e}")
            return None
    
    def run(self):
        """运行合并程序"""
        print("=" * 50)
        print("Excel表格合并工具")
        print("=" * 50)
        
        # 1. 扫描Excel文件
        if not self.scan_excel_files():
            return
        
        # 2. 获取sheet列表
        if not self.get_sheet_names():
            return
        
        # 3. 让用户选择sheet
        target_sheet = self.select_sheet()
        if not target_sheet:
            return
        
        # 4. 让用户选择表头行数
        if not self.select_header_rows():
            return
        
        # 5. 合并数据
        header_rows, merged_data, file_order = self.merge_sheets(target_sheet)
        if merged_data is None:
            return
        
        # 6. 创建输出文件
        output_path = self.create_output_file(header_rows, merged_data, target_sheet, file_order)
        if output_path:
            print(f"\n✅ 合并完成！")
        else:
            print(f"\n❌ 合并失败！")

def main():
    """主函数"""
    try:
        # 创建合并器实例
        merger = ExcelMerger(".")
        
        # 运行合并程序
        merger.run()
        
    except KeyboardInterrupt:
        print("\n\n程序已被用户中断")
    except Exception as e:
        print(f"\n程序运行出错: {e}")
    finally:
        input("\n按回车键退出...")

if __name__ == "__main__":
    main()