#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown流水记录转Excel表格工具
高度自由化和通用化的图形化CLI工具

Author: Claude Code Assistant
Date: 2025-09-20
"""

import os
import re
import json
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
import pandas as pd
from pathlib import Path
import argparse
from typing import Dict, List, Any, Optional
import yaml

class MarkdownToExcelConverter:
    """Markdown转Excel转换器"""

    def __init__(self):
        self.config = self.load_default_config()
        self.data = []
        self.field_mappings = {}

    def load_default_config(self) -> Dict[str, Any]:
        """加载默认配置"""
        return {
            "parsing_rules": {
                "date_pattern": r"## (\d{4}-\d{2}-\d{2})",
                "entry_pattern": r"### (.+?)(?=###|##|$)",
                "header_name_extraction": True,  # 启用从标题行提取姓名
                "field_patterns": {
                    "姓名": r"姓名[：:]\s*([^；;]+)",
                    "交易笔数": r"交易笔数[：:]\s*([^；;]+)",
                    "交易金额": r"交易金额[：:]\s*([^；;]+)",
                    "交易日期": r"交易日期[：:]\s*([^；;]+)",
                    "交易时间": r"交易时间[：:]\s*([^；;]+)",
                    "交易流水号": r"交易流水号[/订单号/交易单号/流水号]*[：:]\s*([^；;]+)",
                    "支付方式": r"支付方式[：:]\s*([^；;]+)",
                    "商户": r"商户[：:]\s*([^；;]+)",
                    "备注": r"备注[：:]\s*([^；;]+)"
                },
                "header_name_patterns": [
                    r"^客户姓名[：:]\s*(.+?)(?:\s|$)",
                    r"^姓名[：:]\s*(.+?)(?:\s|$)",
                    r"^(.+?)(?:\s|$)"
                ]
            },
            "output_settings": {
                "excel_filename": "转换结果_{timestamp}.xlsx",
                "sheet_name": "交易记录",
                "include_summary": True,
                "auto_resize_columns": True
            },
            "validation": {
                "required_fields": ["姓名", "交易金额"],
                "amount_validation": True,
                "date_validation": True
            }
        }

    def save_config(self, config_path: str):
        """保存配置到文件"""
        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.dump(self.config, f, allow_unicode=True, default_flow_style=False)

    def load_config(self, config_path: str):
        """从文件加载配置"""
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                self.config = yaml.safe_load(f)

    def parse_markdown_content(self, content: str) -> List[Dict[str, str]]:
        """解析Markdown内容"""
        data = []
        current_date = None

        # 按日期分段
        date_pattern = self.config["parsing_rules"]["date_pattern"]
        entry_pattern = self.config["parsing_rules"]["entry_pattern"]

        # 找到所有日期
        date_matches = re.finditer(date_pattern, content, re.MULTILINE)
        sections = []

        last_end = 0
        for match in date_matches:
            if last_end < match.start():
                sections.append((None, content[last_end:match.start()]))

            # 找到这个日期下一个日期的位置
            next_date = re.search(date_pattern, content[match.end():])
            if next_date:
                section_end = match.end() + next_date.start()
            else:
                section_end = len(content)

            sections.append((match.group(1), content[match.start():section_end]))
            last_end = section_end

        # 解析每个日期段
        for date, section in sections:
            if date is None:
                continue

            current_date = date

            # 按条目分割
            entries = re.split(r'### ', section)[1:]  # 第一个是标题部分，跳过

            for entry in entries:
                if not entry.strip():
                    continue

                record = {"日期": current_date}

                # 尝试从标题行提取客户姓名
                header_name = self.extract_name_from_header(entry)
                if header_name:
                    record["姓名"] = header_name

                # 提取各字段
                for field_name, pattern in self.config["parsing_rules"]["field_patterns"].items():
                    match = re.search(pattern, entry)
                    if match:
                        value = match.group(1).strip()
                        # 清理数据
                        value = value.rstrip('；;。.')
                        record[field_name] = value
                    else:
                        # 如果正常字段匹配失败，但没有从标题提取姓名，尝试其他方式
                        if field_name == "姓名" and not record.get("姓名"):
                            record[field_name] = ""
                        else:
                            record[field_name] = record.get(field_name, "")

                # 数据清理和验证
                if self.validate_record(record):
                    data.append(record)

        return data

    def extract_name_from_header(self, entry: str) -> Optional[str]:
        """从条目标题行提取客户姓名"""
        # 检查是否启用标题行姓名提取
        if not self.config["parsing_rules"].get("header_name_extraction", True):
            return None

        # 获取第一行作为标题行
        lines = entry.strip().split('\n')
        if not lines:
            return None

        first_line = lines[0].strip()

        # 使用配置文件中的标题匹配模式
        header_patterns = self.config["parsing_rules"].get("header_name_patterns", [
            r"^客户姓名[：:]\s*(.+?)(?:\s|$)",
            r"^姓名[：:]\s*(.+?)(?:\s|$)",
            r"^(.+?)(?:\s|$)"
        ])

        for pattern in header_patterns:
            try:
                match = re.search(pattern, first_line)
                if match:
                    name = match.group(1).strip()
                    # 清理姓名数据
                    name = name.rstrip('；;。.')
                    # 简单验证：姓名长度应该合理（1-10个字符）且不包含数字和特殊字符
                    if 1 <= len(name) <= 10 and not re.search(r'[\d\(\)\[\]/\\\-_=+]', name):
                        return name
            except re.error:
                # 如果正则表达式有错误，跳过这个模式
                continue

        return None

    def validate_record(self, record: Dict[str, str]) -> bool:
        """验证记录数据"""
        if not self.config["validation"]["required_fields"]:
            return True

        for field in self.config["validation"]["required_fields"]:
            if not record.get(field, "").strip():
                return False

        # 金额验证
        if self.config["validation"]["amount_validation"]:
            amount = record.get("交易金额", "")
            if amount and not re.search(r'\d+\.?\d*', amount):
                return False

        return True

    def clean_amount(self, amount_str: str) -> float:
        """清理金额字符串，返回数值"""
        if not amount_str:
            return 0.0

        # 提取数字
        match = re.search(r'(\d+\.?\d*)', amount_str.replace(',', ''))
        if match:
            return float(match.group(1))
        return 0.0

    def export_to_excel(self, data: List[Dict[str, str]], output_path: str):
        """导出到Excel"""
        if not data:
            raise ValueError("没有数据可导出")

        df = pd.DataFrame(data)

        # 处理金额列
        if "交易金额" in df.columns:
            df["金额数值"] = df["交易金额"].apply(self.clean_amount)

        # 创建Excel写入器
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 主数据表
            sheet_name = self.config["output_settings"]["sheet_name"]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # 自动调整列宽
            if self.config["output_settings"]["auto_resize_columns"]:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            # 创建汇总表
            if self.config["output_settings"]["include_summary"] and "金额数值" in df.columns:
                summary_data = []

                # 按日期汇总
                if "日期" in df.columns:
                    daily_summary = df.groupby("日期")["金额数值"].agg(['count', 'sum']).reset_index()
                    daily_summary.columns = ["日期", "交易笔数", "总金额"]
                    summary_data.append(("按日期汇总", daily_summary))

                # 按支付方式汇总
                if "支付方式" in df.columns:
                    payment_summary = df.groupby("支付方式")["金额数值"].agg(['count', 'sum']).reset_index()
                    payment_summary.columns = ["支付方式", "交易笔数", "总金额"]
                    summary_data.append(("按支付方式汇总", payment_summary))

                # 按客户汇总
                if "姓名" in df.columns:
                    customer_summary = df.groupby("姓名")["金额数值"].agg(['count', 'sum']).reset_index()
                    customer_summary.columns = ["客户姓名", "交易笔数", "总金额"]
                    summary_data.append(("按客户汇总", customer_summary))

                # 写入汇总表
                for sheet_name, summary_df in summary_data:
                    summary_df.to_excel(writer, sheet_name=sheet_name, index=False)


class MarkdownToExcelGUI:
    """图形化用户界面"""

    def __init__(self):
        self.converter = MarkdownToExcelConverter()
        self.setup_gui()

    def setup_gui(self):
        """设置GUI界面"""
        self.root = tk.Tk()
        self.root.title("Markdown流水转Excel工具 v1.0")
        self.root.geometry("900x700")

        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置行列权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # 文件选择部分
        ttk.Label(main_frame, text="输入文件:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)

        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(0, weight=1)

        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        self.file_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))

        ttk.Button(file_frame, text="浏览", command=self.browse_file).grid(row=0, column=1)

        # 输出设置
        ttk.Label(main_frame, text="输出设置:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=(20, 5))

        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(20, 5))
        output_frame.columnconfigure(0, weight=1)

        self.output_path_var = tk.StringVar()
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path_var, width=50)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))

        ttk.Button(output_frame, text="选择目录", command=self.browse_output_dir).grid(row=0, column=1)

        # 配置选项
        config_frame = ttk.LabelFrame(main_frame, text="配置选项", padding="10")
        config_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        config_frame.columnconfigure(1, weight=1)

        # 包含汇总表选项
        self.include_summary_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(config_frame, text="包含汇总表",
                       variable=self.include_summary_var).grid(row=0, column=0, sticky=tk.W)

        # 自动调整列宽选项
        self.auto_resize_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(config_frame, text="自动调整列宽",
                       variable=self.auto_resize_var).grid(row=0, column=1, sticky=tk.W)

        # 数据验证选项
        self.validate_data_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(config_frame, text="启用数据验证",
                       variable=self.validate_data_var).grid(row=1, column=0, sticky=tk.W)

        # 自定义字段映射
        ttk.Label(config_frame, text="自定义字段正则 (可选):").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))

        self.custom_fields_text = scrolledtext.ScrolledText(config_frame, height=6, width=60)
        self.custom_fields_text.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # 插入默认配置示例
        example_config = """# 字段正则表达式配置示例 (YAML格式):
# 姓名: "姓名[：:]\\s*([^；;]+)"
# 交易金额: "交易金额[：:]\\s*([^；;]+)"
# 自定义字段: "自定义字段[：:]\\s*([^；;]+)"
"""
        self.custom_fields_text.insert(tk.END, example_config)

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)

        ttk.Button(button_frame, text="预览数据", command=self.preview_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="开始转换", command=self.convert_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="加载配置", command=self.load_config).pack(side=tk.LEFT, padx=5)

        # 状态和日志区域
        status_frame = ttk.LabelFrame(main_frame, text="状态信息", padding="10")
        status_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        status_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)

        self.status_text = scrolledtext.ScrolledText(status_frame, height=10, width=80)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.log("工具已启动，请选择要转换的Markdown文件...")

    def log(self, message: str):
        """添加日志信息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update()

    def browse_file(self):
        """浏览文件"""
        filename = filedialog.askopenfilename(
            title="选择Markdown文件",
            filetypes=[("Markdown files", "*.md"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
            self.log(f"已选择文件: {filename}")

            # 自动设置输出路径
            input_path = Path(filename)
            output_dir = input_path.parent
            self.output_path_var.set(str(output_dir))

    def browse_output_dir(self):
        """浏览输出目录"""
        dirname = filedialog.askdirectory(title="选择输出目录")
        if dirname:
            self.output_path_var.set(dirname)
            self.log(f"已选择输出目录: {dirname}")

    def update_converter_config(self):
        """更新转换器配置"""
        # 更新基本设置
        self.converter.config["output_settings"]["include_summary"] = self.include_summary_var.get()
        self.converter.config["output_settings"]["auto_resize_columns"] = self.auto_resize_var.get()

        # 处理自定义字段配置
        custom_config = self.custom_fields_text.get("1.0", tk.END).strip()
        if custom_config and not custom_config.startswith("#"):
            try:
                # 尝试解析YAML格式的字段配置
                custom_fields = yaml.safe_load(custom_config)
                if isinstance(custom_fields, dict):
                    self.converter.config["parsing_rules"]["field_patterns"].update(custom_fields)
                    self.log("已应用自定义字段配置")
            except Exception as e:
                self.log(f"自定义字段配置解析失败: {e}")

    def preview_data(self):
        """预览数据"""
        if not self.file_path_var.get():
            messagebox.showerror("错误", "请先选择输入文件！")
            return

        try:
            self.progress.start()
            self.update_converter_config()

            with open(self.file_path_var.get(), 'r', encoding='utf-8') as f:
                content = f.read()

            data = self.converter.parse_markdown_content(content)

            if not data:
                self.log("未找到有效数据！")
                return

            # 显示预览窗口
            self.show_preview_window(data)
            self.log(f"数据预览完成，共找到 {len(data)} 条记录")

        except Exception as e:
            self.log(f"预览失败: {e}")
            messagebox.showerror("错误", f"预览失败: {e}")
        finally:
            self.progress.stop()

    def show_preview_window(self, data: List[Dict[str, str]]):
        """显示预览窗口"""
        preview_window = tk.Toplevel(self.root)
        preview_window.title("数据预览")
        preview_window.geometry("800x600")

        # 创建表格
        columns = list(data[0].keys()) if data else []
        tree = ttk.Treeview(preview_window, columns=columns, show='headings', height=20)

        # 设置列标题
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        # 插入数据 (只显示前100条)
        for i, record in enumerate(data[:100]):
            values = [record.get(col, "") for col in columns]
            tree.insert("", tk.END, values=values)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(preview_window, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        if len(data) > 100:
            info_label = ttk.Label(preview_window, text=f"显示前100条记录，总共{len(data)}条")
            info_label.pack(side=tk.BOTTOM)

    def convert_file(self):
        """转换文件"""
        if not self.file_path_var.get():
            messagebox.showerror("错误", "请先选择输入文件！")
            return

        if not self.output_path_var.get():
            messagebox.showerror("错误", "请先选择输出目录！")
            return

        try:
            self.progress.start()
            self.update_converter_config()

            # 读取文件
            self.log("正在读取文件...")
            with open(self.file_path_var.get(), 'r', encoding='utf-8') as f:
                content = f.read()

            # 解析数据
            self.log("正在解析数据...")
            data = self.converter.parse_markdown_content(content)

            if not data:
                self.log("未找到有效数据！")
                messagebox.showerror("错误", "未找到有效数据！")
                return

            # 生成输出文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = self.converter.config["output_settings"]["excel_filename"].format(timestamp=timestamp)
            output_path = os.path.join(self.output_path_var.get(), filename)

            # 导出Excel
            self.log("正在导出Excel...")
            self.converter.export_to_excel(data, output_path)

            self.log(f"转换完成！")
            self.log(f"输出文件: {output_path}")
            self.log(f"共转换 {len(data)} 条记录")

            # 询问是否打开文件
            if messagebox.askyesno("完成", f"转换完成！\n\n输出文件: {output_path}\n共转换 {len(data)} 条记录\n\n是否打开文件？"):
                os.startfile(output_path)

        except Exception as e:
            self.log(f"转换失败: {e}")
            messagebox.showerror("错误", f"转换失败: {e}")
        finally:
            self.progress.stop()

    def save_config(self):
        """保存配置"""
        filename = filedialog.asksaveasfilename(
            title="保存配置文件",
            defaultextension=".yaml",
            filetypes=[("YAML files", "*.yaml"), ("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                self.update_converter_config()
                self.converter.save_config(filename)
                self.log(f"配置已保存到: {filename}")
                messagebox.showinfo("成功", "配置保存成功！")
            except Exception as e:
                self.log(f"保存配置失败: {e}")
                messagebox.showerror("错误", f"保存配置失败: {e}")

    def load_config(self):
        """加载配置"""
        filename = filedialog.askopenfilename(
            title="加载配置文件",
            filetypes=[("YAML files", "*.yaml"), ("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                self.converter.load_config(filename)

                # 更新GUI设置
                self.include_summary_var.set(self.converter.config["output_settings"]["include_summary"])
                self.auto_resize_var.set(self.converter.config["output_settings"]["auto_resize_columns"])

                self.log(f"配置已从文件加载: {filename}")
                messagebox.showinfo("成功", "配置加载成功！")
            except Exception as e:
                self.log(f"加载配置失败: {e}")
                messagebox.showerror("错误", f"加载配置失败: {e}")

    def run(self):
        """运行GUI"""
        self.root.mainloop()


def create_cli_parser():
    """创建命令行参数解析器"""
    parser = argparse.ArgumentParser(description="Markdown流水转Excel工具")
    parser.add_argument("input", nargs="?", help="输入的Markdown文件路径")
    parser.add_argument("-o", "--output", help="输出Excel文件路径")
    parser.add_argument("-c", "--config", help="配置文件路径")
    parser.add_argument("--no-gui", action="store_true", help="不使用GUI界面")
    parser.add_argument("--no-summary", action="store_true", help="不生成汇总表")
    return parser


def main():
    """主函数"""
    parser = create_cli_parser()
    args = parser.parse_args()

    if args.no_gui and args.input:
        # 命令行模式
        converter = MarkdownToExcelConverter()

        if args.config:
            converter.load_config(args.config)

        if args.no_summary:
            converter.config["output_settings"]["include_summary"] = False

        try:
            with open(args.input, 'r', encoding='utf-8') as f:
                content = f.read()

            data = converter.parse_markdown_content(content)

            if not data:
                print("错误: 未找到有效数据!")
                sys.exit(1)

            if args.output:
                output_path = args.output
            else:
                input_path = Path(args.input)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = input_path.parent / f"转换结果_{timestamp}.xlsx"

            converter.export_to_excel(data, str(output_path))
            print(f"转换完成! 输出文件: {output_path}")
            print(f"共转换 {len(data)} 条记录")

        except Exception as e:
            print(f"错误: {e}")
            sys.exit(1)
    else:
        # GUI模式
        try:
            app = MarkdownToExcelGUI()
            if args.input:
                app.file_path_var.set(args.input)
            app.run()
        except ImportError as e:
            print("错误: 缺少GUI依赖库")
            print("请安装tkinter: pip install tk")
            sys.exit(1)


if __name__ == "__main__":
    main()