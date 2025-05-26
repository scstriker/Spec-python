import sys
import os
import pandas as pd
import random
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                            QSpinBox, QTableWidget, QTableWidgetItem, QHeaderView,
                            QMessageBox, QLineEdit, QComboBox, QGroupBox)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QIcon, QPalette, QColor

class ExpertLotteryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.experts_data = None
        self.extracted_experts = None
        self.research_fields = []
        self.organizations = []
        self.apply_win98_style()
        self.init_ui()
        
    def apply_win98_style(self):
        """应用Windows 98风格的样式"""
        # 设置应用程序样式表
        win98_style = """
        QMainWindow, QDialog {
            background-color: #c0c0c0;
        }
        QWidget {
            background-color: #c0c0c0;
            color: #000000;
        }
        QPushButton {
            background-color: #c0c0c0;
            border: 2px solid #808080;
            border-top-color: #ffffff;
            border-left-color: #ffffff;
            border-right-color: #404040;
            border-bottom-color: #404040;
            padding: 3px;
            min-width: 80px;
            color: #000000;
        }
        QPushButton:hover {
            background-color: #d0d0d0;
        }
        QPushButton:pressed {
            border: 2px solid #808080;
            border-top-color: #404040;
            border-left-color: #404040;
            border-right-color: #ffffff;
            border-bottom-color: #ffffff;
            background-color: #b0b0b0;
        }
        QLineEdit, QSpinBox, QComboBox {
            background-color: #ffffff;
            border: 1px solid #808080;
            border-top-color: #404040;
            border-left-color: #404040;
            border-right-color: #ffffff;
            border-bottom-color: #ffffff;
            padding: 2px;
            color: #000000;
        }
        QComboBox::drop-down {
            border: 0px;
            width: 20px;
        }
        QComboBox::down-arrow {
            image: url(none);
            width: 16px;
            height: 16px;
            background-color: #c0c0c0;
            border: 1px solid #808080;
            border-top-color: #ffffff;
            border-left-color: #ffffff;
            border-right-color: #404040;
            border-bottom-color: #404040;
        }
        QTableWidget {
            background-color: #ffffff;
            alternate-background-color: #f0f0f0;
            gridline-color: #808080;
            border: 2px solid #808080;
            border-top-color: #404040;
            border-left-color: #404040;
            border-right-color: #ffffff;
            border-bottom-color: #ffffff;
            selection-background-color: #000080;
            selection-color: #ffffff;
        }
        QHeaderView::section {
            background-color: #c0c0c0;
            border: 2px solid #808080;
            border-top-color: #ffffff;
            border-left-color: #ffffff;
            border-right-color: #404040;
            border-bottom-color: #404040;
            padding: 2px;
            color: #000000;
        }
        QGroupBox {
            border: 2px solid #808080;
            border-top-color: #404040;
            border-left-color: #404040;
            border-right-color: #ffffff;
            border-bottom-color: #ffffff;
            margin-top: 10px;
            font-weight: bold;
            color: #000000;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            left: 10px;
            padding: 0 3px;
        }
        """
        self.setStyleSheet(win98_style)
        
        # 设置应用程序调色板
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(192, 192, 192))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(255, 255, 220))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.Text, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.Button, QColor(192, 192, 192))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(0, 0, 128))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
        QApplication.instance().setPalette(palette)
        
    def init_ui(self):
        # 设置窗口属性
        self.setWindowTitle("专家抽奖系统")
        self.setMinimumSize(1000, 650)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 创建标题标签
        title_label = QLabel("专家抽奖系统")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        main_layout.addWidget(title_label)
        
        # 文件选择区域
        file_section = QWidget()
        file_layout = QHBoxLayout(file_section)
        
        file_label = QLabel("专家名单文件:")
        self.file_path_display = QLineEdit()
        self.file_path_display.setReadOnly(True)
        self.file_path_display.setPlaceholderText("请选择Excel文件...")
        
        browse_button = QPushButton("浏览...")
        browse_button.clicked.connect(self.browse_file)
        
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_path_display, 1)
        file_layout.addWidget(browse_button)
        
        main_layout.addWidget(file_section)
        
        # 筛选设置区域
        filter_section = QWidget()
        filter_layout = QHBoxLayout(filter_section)
        
        # 研究领域选择框
        field_group = QGroupBox("研究领域筛选")
        field_group_layout = QVBoxLayout(field_group)
        
        self.field_combo = QComboBox()
        self.field_combo.addItem("全部领域")
        self.field_combo.setMinimumWidth(150)
        
        field_group_layout.addWidget(self.field_combo)
        filter_layout.addWidget(field_group)
        
        # 回避单位选择框
        avoid_group = QGroupBox("回避单位")
        avoid_group_layout = QVBoxLayout(avoid_group)
        
        self.avoid_combo = QComboBox()
        self.avoid_combo.addItem("不回避任何单位")
        self.avoid_combo.setMinimumWidth(150)
        
        avoid_group_layout.addWidget(self.avoid_combo)
        filter_layout.addWidget(avoid_group)
        
        main_layout.addWidget(filter_section)
        
        # 抽取设置区域
        extraction_section = QWidget()
        extraction_layout = QHBoxLayout(extraction_section)
        
        count_label = QLabel("抽取数量:")
        self.count_spinbox = QSpinBox()
        self.count_spinbox.setMinimum(1)
        self.count_spinbox.setMaximum(9999)
        self.count_spinbox.setValue(5)
        
        extract_button = QPushButton("开始抽奖")
        extract_button.clicked.connect(self.extract_experts)
        
        save_button = QPushButton("保存结果")
        save_button.clicked.connect(self.save_results)
        
        extraction_layout.addWidget(count_label)
        extraction_layout.addWidget(self.count_spinbox)
        extraction_layout.addStretch(1)
        extraction_layout.addWidget(extract_button)
        extraction_layout.addWidget(save_button)
        
        main_layout.addWidget(extraction_section)
        
        # 创建表格
        self.create_table()
        main_layout.addWidget(self.table)
        
        # 状态标签
        self.status_label = QLabel("请选择专家名单Excel文件开始抽奖")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
        
    def create_table(self):
        """创建表格以显示专家信息"""
        self.table = QTableWidget()
        self.table.setColumnCount(0)
        self.table.setRowCount(0)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setAlternatingRowColors(True)
        
    def browse_file(self):
        """打开文件对话框选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择专家名单Excel文件", "", "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.file_path_display.setText(file_path)
            self.load_experts_file(file_path)
    
    def load_experts_file(self, file_path):
        """从Excel文件加载专家名单"""
        try:
            self.experts_data = pd.read_excel(file_path)
            self.status_label.setText(f"成功加载了 {len(self.experts_data)} 位专家信息")
            
            # 更新最大可抽取数量
            self.count_spinbox.setMaximum(len(self.experts_data))
            
            # 清空表格
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            
            # 更新研究领域和单位下拉框
            self.update_research_fields()
            self.update_organizations()
            
        except FileNotFoundError:
            QMessageBox.critical(self, "错误", f"文件 '{file_path}' 未找到。请检查文件路径是否正确。")
            self.status_label.setText("文件加载失败")
            self.experts_data = None
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取文件时发生错误: {e}")
            self.status_label.setText("文件加载失败")
            self.experts_data = None
    
    def update_research_fields(self):
        """更新研究领域下拉框"""
        if self.experts_data is not None and '研究领域' in self.experts_data.columns:
            # 保存当前选中项
            current_text = self.field_combo.currentText()
            
            # 清空下拉框
            self.field_combo.clear()
            self.field_combo.addItem("全部领域")
            
            # 获取所有唯一的研究领域
            self.research_fields = self.experts_data['研究领域'].unique().tolist()
            
            # 添加到下拉框
            for field in self.research_fields:
                self.field_combo.addItem(field)
                
            # 尝试恢复之前的选择
            index = self.field_combo.findText(current_text)
            if index >= 0:
                self.field_combo.setCurrentIndex(index)
                
    def update_organizations(self):
        """更新单位下拉框"""
        if self.experts_data is not None and '单位' in self.experts_data.columns:
            # 保存当前选中项
            current_text = self.avoid_combo.currentText()
            
            # 清空下拉框
            self.avoid_combo.clear()
            self.avoid_combo.addItem("不回避任何单位")
            
            # 获取所有唯一的单位
            self.organizations = self.experts_data['单位'].unique().tolist()
            
            # 添加到下拉框
            for org in self.organizations:
                self.avoid_combo.addItem(org)
                
            # 尝试恢复之前的选择
            index = self.avoid_combo.findText(current_text)
            if index >= 0:
                self.avoid_combo.setCurrentIndex(index)
    
    def extract_experts(self):
        """随机抽取专家"""
        if self.experts_data is None or self.experts_data.empty:
            QMessageBox.warning(self, "警告", "请先选择并加载专家名单文件")
            return
        
        num_to_extract = self.count_spinbox.value()
        
        if num_to_extract <= 0:
            QMessageBox.warning(self, "警告", "抽取的专家数量必须大于0")
            return
            
        # 筛选数据
        filtered_data = self.filter_experts_data()
        
        if filtered_data.empty:
            QMessageBox.warning(self, "警告", "根据当前筛选条件，没有找到符合条件的专家")
            return
            
        num_available_experts = len(filtered_data)
        
        if num_to_extract > num_available_experts:
            QMessageBox.warning(
                self, 
                "警告", 
                f"要抽取的专家数量 ({num_to_extract}) 大于可用专家总数 ({num_available_experts})。\n将返回所有可用专家。"
            )
            self.extracted_experts = filtered_data.copy()
        else:
            # 使用pandas的sample方法进行随机抽样
            self.extracted_experts = filtered_data.sample(
                n=num_to_extract, 
                random_state=random.randint(1, 1000)
            )
        
        # 显示抽取结果
        self.display_experts(self.extracted_experts)
        
        # 更新状态栏信息
        self.update_status_info()
        
    def filter_experts_data(self):
        """根据筛选条件过滤专家数据"""
        if self.experts_data is None or self.experts_data.empty:
            return pd.DataFrame()
            
        # 复制数据
        filtered_data = self.experts_data.copy()
        
        # 检查是否需要按研究领域筛选
        selected_field = self.field_combo.currentText()
        if selected_field != "全部领域" and '研究领域' in self.experts_data.columns:
            filtered_data = filtered_data[filtered_data['研究领域'] == selected_field]
        
        # 检查是否需要回避某个单位
        avoided_org = self.avoid_combo.currentText()
        if avoided_org != "不回避任何单位" and '单位' in self.experts_data.columns:
            filtered_data = filtered_data[filtered_data['单位'] != avoided_org]
            
        return filtered_data
        
    def update_status_info(self):
        """更新状态栏信息"""
        if self.extracted_experts is None or self.extracted_experts.empty:
            self.status_label.setText("未找到符合条件的专家")
            return
            
        # 获取筛选信息
        field_info = ""
        avoid_info = ""
        
        selected_field = self.field_combo.currentText()
        if selected_field != "全部领域":
            field_info = f"[{selected_field}] "
            
        avoided_org = self.avoid_combo.currentText()
        if avoided_org != "不回避任何单位":
            avoid_info = f"[回避:{avoided_org}] "
            
        self.status_label.setText(f"已随机抽取 {field_info}{avoid_info}{len(self.extracted_experts)} 位专家")
    
    def display_experts(self, experts_df):
        """在表格中显示专家信息"""
        if experts_df is None or experts_df.empty:
            return
            
        # 设置表格列数
        columns = experts_df.columns
        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)
        
        # 设置表格行数
        rows = len(experts_df)
        self.table.setRowCount(rows)
        
        # 填充表格数据
        for row in range(rows):
            for col in range(len(columns)):
                value = str(experts_df.iloc[row, col])
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(row, col, item)
    
    def save_results(self):
        """保存抽取结果到新的Excel文件"""
        if self.extracted_experts is None or self.extracted_experts.empty:
            QMessageBox.warning(self, "警告", "没有可以保存的抽奖结果，请先进行抽奖")
            return
            
        # 生成默认文件名
        selected_field = self.field_combo.currentText()
        avoided_org = self.avoid_combo.currentText()
        
        default_filename = "抽奖结果.xlsx"
        
        filename_parts = []
        if selected_field != "全部领域":
            filename_parts.append(selected_field)
            
        if avoided_org != "不回避任何单位":
            filename_parts.append(f"回避{avoided_org}")
            
        if filename_parts:
            default_filename = f"{'_'.join(filename_parts)}-抽奖结果.xlsx"
            
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存抽奖结果", default_filename, "Excel Files (*.xlsx)"
        )
        
        if file_path:
            try:
                self.extracted_experts.to_excel(file_path, index=False)
                QMessageBox.information(self, "成功", f"抽奖结果已成功保存到 '{file_path}'")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存文件时发生错误: {e}")


def main():
    app = QApplication(sys.argv)
    window = ExpertLotteryApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main() 