import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QLabel, QLineEdit, QFileDialog, QMessageBox,
    QListWidget, QListWidgetItem, QInputDialog
)
from PyQt5.QtCore import QDateTime, Qt
from PyQt5.QtGui import QDesktopServices, QDragEnterEvent, QDropEvent
from PyQt5.QtCore import QUrl

class ExcelCleaner(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("国考系统_批量数据清洗工具")
        self.resize(1000, 700)

        self.input_files = []   # 存储导入文件路径
        self.output_files = []  # 存储导出文件名，可修改
        self.export_folder = None

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 文件操作按钮
        file_btn_layout = QHBoxLayout()
        self.import_btn = QPushButton("【1】导入Excel文件（可批量拖拽至窗口）")
        self.import_btn.clicked.connect(self.import_files)
        file_btn_layout.addWidget(self.import_btn)

        self.select_export_folder_btn = QPushButton("【2】选择导出文件夹")
        self.select_export_folder_btn.clicked.connect(self.select_export_folder)
        file_btn_layout.addWidget(self.select_export_folder_btn)

        self.edit_output_name_btn = QPushButton("修改导出文件名（可选）")
        self.edit_output_name_btn.clicked.connect(self.edit_output_name)
        file_btn_layout.addWidget(self.edit_output_name_btn)

        self.open_input_folder_btn = QPushButton("打开导入文件夹（可选）")
        self.open_input_folder_btn.clicked.connect(self.open_input_folder)
        file_btn_layout.addWidget(self.open_input_folder_btn)

        self.open_output_folder_btn = QPushButton("打开导出文件夹（可选）")
        self.open_output_folder_btn.clicked.connect(self.open_output_folder)
        file_btn_layout.addWidget(self.open_output_folder_btn)

        layout.addLayout(file_btn_layout)

        # 文件列表
        self.file_list_widget = QListWidget()
        self.file_list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        layout.addWidget(self.file_list_widget)

        # 批量转换按钮
        self.run_btn = QPushButton("【3】开始批量转换并导出")
        self.run_btn.clicked.connect(self.convert_all)
        layout.addWidget(self.run_btn)

        # 日志窗口
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        # 开发者信息
        dev_label = QLabel("开发者: 欧星星 | v1.0")
        dev_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        dev_label.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(dev_label)

        # 支持拖拽
        self.setAcceptDrops(True)

    def log(self, message: str, error: bool = False):
        timestamp = QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss")
        msg = f"[{timestamp}] {message}"
        if error:
            msg = f"<span style='color:red;'>{msg}</span>"
        self.log_text.append(msg)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    # 拖拽事件
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith((".xls", ".xlsx")):
                self.add_file(path)

    # 导入文件
    def import_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        for f in files:
            self.add_file(f)

    # 添加文件到列表
    def add_file(self, file_path):
        if file_path in self.input_files:
            return
        self.input_files.append(file_path)
        base_name = os.path.basename(file_path)
        default_output = f"清洗_{base_name}"
        self.output_files.append(default_output)
        item = QListWidgetItem(f"{base_name} -> {default_output}")
        self.file_list_widget.addItem(item)
        self.log(f"成功导入文件: {base_name}")

    # 修改选中文件的导出文件名
    def edit_output_name(self):
        selected_items = self.file_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return
        for item in selected_items:
            idx = self.file_list_widget.row(item)
            current_name = self.output_files[idx]
            new_name, ok = QInputDialog.getText(self, "修改导出文件名", "请输入新的导出文件名:", text=current_name)
            if ok and new_name:
                self.output_files[idx] = new_name
                base_name = os.path.basename(self.input_files[idx])
                item.setText(f"{base_name} -> {new_name}")

    # 选择统一导出文件夹
    def select_export_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择导出文件夹")
        if folder:
            self.export_folder = folder
            self.log(f"已选择统一导出文件夹: {folder}")

    # 打开导入文件夹
    def open_input_folder(self):
        if not self.input_files:
            QMessageBox.warning(self, "提示", "没有导入文件")
            return
        folder = os.path.dirname(self.input_files[0])
        QDesktopServices.openUrl(QUrl.fromLocalFile(folder))

    # 打开导出文件夹
    def open_output_folder(self):
        if not self.export_folder:
            QMessageBox.warning(self, "提示", "请先选择导出文件夹")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.export_folder))

    # 批量转换
    def convert_all(self):
        if not self.input_files:
            QMessageBox.warning(self, "提示", "请先导入文件！")
            return
        if not self.export_folder:
            QMessageBox.warning(self, "提示", "请先选择统一导出文件夹！")
            return

        for idx, file_path in enumerate(self.input_files):
            try:
                self.log(f"开始处理文件: {os.path.basename(file_path)}")
                df = pd.read_excel(file_path)
                df.columns = [str(c).strip() for c in df.columns]

                month_cols = [c for c in df.columns if c != "科室名称" and c != "合计"]
                records = []
                for dept_idx, row in df.iterrows():
                    dept_name = row["科室名称"]
                    self.log(f"  处理科室: {dept_name}")
                    for month in month_cols:
                        val = row[month]
                        self.log(f"    月份 {month} 出院人数: {val}")
                        records.append({
                            "科室名称": dept_name,
                            "日期": month,
                            "出院人数": val
                        })

                melted = pd.DataFrame(records)
                dept_order = {name: i for i, name in enumerate(df["科室名称"].tolist())}
                melted["日期_dt"] = pd.to_datetime(melted["日期"], format="%Y-%m")
                melted["dept_order"] = melted["科室名称"].map(dept_order)
                melted = melted.sort_values(by=["dept_order", "日期_dt"]).reset_index(drop=True)
                melted.drop(columns=["dept_order", "日期_dt"], inplace=True)
                melted["年份"] = pd.to_datetime(melted["日期"], format="%Y-%m").dt.year.astype(str) + "年"
                melted["月份"] = pd.to_datetime(melted["日期"], format="%Y-%m").dt.month.apply(lambda x: f"{x:02d}月")
                melted.insert(0, "序号", range(1, len(melted)+1))
                melted.rename(columns={"科室名称": "科室"}, inplace=True)
                melted = melted[["序号", "科室", "日期", "年份", "月份", "出院人数"]]

                # 导出到统一文件夹
                output_path = os.path.join(self.export_folder, self.output_files[idx])
                melted.to_excel(output_path, index=False)
                self.log(f"成功导出文件: {self.output_files[idx]}")

            except Exception as e:
                self.log(f"文件处理出错: {file_path} 错误信息: {e}", error=True)
                QMessageBox.critical(self, "错误", f"文件处理出错: {file_path}\n错误信息: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelCleaner()
    window.show()
    sys.exit(app.exec())
