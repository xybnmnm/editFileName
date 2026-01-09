#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量重命名工具 - 最终修正版
修正逻辑：根据Excel映射表重命名目标文件夹中的文件
优化界面：简化输入项，增大输出框
"""

import os
import sys
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import traceback
import subprocess

class BatchRenamer:
    """批量重命名工具主类 - 最终修正版"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("批量重命名工具")

        # 设置窗口最小尺寸
        self.root.minsize(1000, 800)

        # 设置样式
        style = ttk.Style()
        if sys.platform == "darwin":  # macOS
            style.theme_use('aqua')
        elif sys.platform == "win32":  # Windows
            style.theme_use('vista')
        else:  # Linux
            style.theme_use('clam')

        # 配置网格权重
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.create_widgets()

    def create_widgets(self):
        """创建界面组件"""

        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # 创建标签页
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 模式1：移除特定字符串
        frame1 = ttk.Frame(notebook, padding="10")
        notebook.add(frame1, text='模式1: 移除指定字符串')

        # 模式2：根据Excel映射重命名
        frame2 = ttk.Frame(notebook, padding="10")
        notebook.add(frame2, text='模式2: Excel映射重命名')

        self.create_mode1_ui(frame1)
        self.create_mode2_ui(frame2)

    def create_mode1_ui(self, parent):
        """创建模式1的界面"""
        # 配置父框架的网格权重
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        # 标题区域
        title_frame = ttk.Frame(parent)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        title_frame.grid_columnconfigure(0, weight=1)

        title_label = ttk.Label(title_frame, text="模式1：移除文件名中的指定字符串",
                                font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, sticky=tk.W)

        desc_label = ttk.Label(title_frame,
                               text="例如：将 25142000000109317099_01.jpg 重命名为 25142000000109317099.jpg\n"
                                    "只需输入要移除的字符串（如 '_01'）并选择文件夹",
                               wraplength=600)
        desc_label.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))

        # 设置区域
        settings_frame = ttk.LabelFrame(parent, text="设置", padding="15")
        settings_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        settings_frame.grid_columnconfigure(1, weight=1)

        # 文件夹选择
        ttk.Label(settings_frame, text="选择文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.folder_path1 = tk.StringVar()
        folder_entry = ttk.Entry(settings_frame, textvariable=self.folder_path1)
        folder_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        ttk.Button(settings_frame, text="浏览...", width=10,
                   command=self.select_folder1).grid(row=0, column=2, padx=5, pady=5)

        # 要移除的字符串
        ttk.Label(settings_frame, text="移除的字符串:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.remove_string = tk.StringVar()
        string_entry = ttk.Entry(settings_frame, textvariable=self.remove_string, width=30)
        string_entry.grid(row=1, column=1, sticky=tk.W, padx=(10, 10), pady=5)
        ttk.Label(settings_frame, text="示例: _01, _abc, 前缀_等").grid(row=1, column=2, sticky=tk.W, pady=5)

        # 按钮区域
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)

        ttk.Button(button_frame, text="预览重命名",
                   command=self.preview_removal).grid(row=0, column=0, padx=2, sticky=tk.EW)
        ttk.Button(button_frame, text="执行重命名",
                   command=self.execute_removal).grid(row=0, column=1, padx=2, sticky=tk.EW)
        ttk.Button(button_frame, text="清空预览",
                   command=self.clear_preview1).grid(row=0, column=2, padx=2, sticky=tk.EW)

        # 预览区域
        preview_frame = ttk.LabelFrame(parent, text="重命名预览", padding="10")
        preview_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        # 创建文本区域和滚动条
        text_frame = ttk.Frame(preview_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)

        self.preview_text1 = tk.Text(text_frame, height=15)
        scrollbar = ttk.Scrollbar(text_frame, command=self.preview_text1.yview)
        self.preview_text1.configure(yscrollcommand=scrollbar.set)

        self.preview_text1.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

    def create_mode2_ui(self, parent):
        """创建模式2的界面 - 最终修正版"""
        # 配置父框架的网格权重
        parent.grid_rowconfigure(5, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        # 标题区域 - 简化
        title_frame = ttk.Frame(parent)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        title_frame.grid_columnconfigure(0, weight=1)

        title_label = ttk.Label(title_frame, text="模式2：Excel映射重命名",
                                font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, sticky=tk.W)

        desc_label = ttk.Label(title_frame,
                               text="根据Excel映射表重命名目标文件夹中的文件\n"
                                    "✓ Excel中：源文件名列（当前文件名） → 目标文件名列（新文件名）\n"
                                    "✓ 目标文件夹：包含需要重命名的文件\n"
                                    "✓ 找不到映射关系的文件保持不变，最后统计更新数量",
                               wraplength=700, justify='left')
        desc_label.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))

        # 文件选择区域 - 简化：只需要目标文件夹和Excel文件
        files_frame = ttk.LabelFrame(parent, text="文件设置", padding="15")
        files_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        files_frame.grid_columnconfigure(1, weight=1)

        # 目标文件夹（包含需要重命名的文件）
        ttk.Label(files_frame, text="目标文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.target_folder_path = tk.StringVar()
        target_entry = ttk.Entry(files_frame, textvariable=self.target_folder_path)
        target_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        ttk.Button(files_frame, text="浏览...", width=10,
                   command=lambda: self.select_folder(self.target_folder_path)).grid(row=0, column=2, padx=5, pady=5)

        # Excel映射文件
        ttk.Label(files_frame, text="Excel映射文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.excel_path = tk.StringVar()
        excel_entry = ttk.Entry(files_frame, textvariable=self.excel_path)
        excel_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        ttk.Button(files_frame, text="浏览...", width=10,
                   command=self.select_excel).grid(row=1, column=2, padx=5, pady=5)

        # Excel列设置区域
        excel_settings_frame = ttk.LabelFrame(parent, text="Excel列设置", padding="15")
        excel_settings_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        excel_settings_frame.grid_columnconfigure(1, weight=1)
        excel_settings_frame.grid_columnconfigure(3, weight=1)

        # 列名设置
        ttk.Label(excel_settings_frame, text="源文件名列（当前文件名）:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.src_col_name = tk.StringVar(value="源文件名")
        ttk.Entry(excel_settings_frame, textvariable=self.src_col_name, width=20).grid(row=0, column=1, sticky=tk.W, padx=(10, 20), pady=5)

        ttk.Label(excel_settings_frame, text="目标文件名列（新文件名）:").grid(row=0, column=2, sticky=tk.W, pady=5)
        self.dst_col_name = tk.StringVar(value="目标文件名")
        ttk.Entry(excel_settings_frame, textvariable=self.dst_col_name, width=20).grid(row=0, column=3, sticky=tk.W, padx=(10, 0), pady=5)

        # 匹配模式和选项
        ttk.Label(excel_settings_frame, text="匹配模式:").grid(row=1, column=0, sticky=tk.W, pady=5)
        match_frame = ttk.Frame(excel_settings_frame)
        match_frame.grid(row=1, column=1, columnspan=3, sticky=tk.W, padx=(10, 0), pady=5)

        self.match_mode = tk.StringVar(value="exact")
        ttk.Radiobutton(match_frame, text="精确匹配", variable=self.match_mode, value="exact").pack(side='left', padx=5)
        ttk.Radiobutton(match_frame, text="包含匹配", variable=self.match_mode, value="contains").pack(side='left', padx=5)

        # 扩展名选项
        ttk.Label(excel_settings_frame, text="扩展名处理:").grid(row=2, column=0, sticky=tk.W, pady=5)
        option_frame = ttk.Frame(excel_settings_frame)
        option_frame.grid(row=2, column=1, columnspan=3, sticky=tk.W, padx=(10, 0), pady=5)

        self.keep_original_extension = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="保留原文件扩展名",
                        variable=self.keep_original_extension).pack(side='left', padx=5)

        # 工具按钮区域
        tool_button_frame = ttk.Frame(parent)
        tool_button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        tool_button_frame.grid_columnconfigure(0, weight=1)
        tool_button_frame.grid_columnconfigure(1, weight=1)
        tool_button_frame.grid_columnconfigure(2, weight=1)

        ttk.Button(tool_button_frame, text="查看Excel列名",
                   command=self.preview_excel_columns).grid(row=0, column=0, padx=2, sticky=tk.EW)
        ttk.Button(tool_button_frame, text="查看Excel完整数据",
                   command=self.preview_full_excel).grid(row=0, column=1, padx=2, sticky=tk.EW)
        ttk.Button(tool_button_frame, text="测试映射关系",
                   command=self.test_mapping).grid(row=0, column=2, padx=2, sticky=tk.EW)

        # 操作按钮区域
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

        # 第一行按钮
        button_row1 = ttk.Frame(button_frame)
        button_row1.pack(fill=tk.X, pady=5)
        button_row1.grid_columnconfigure(0, weight=1)
        button_row1.grid_columnconfigure(1, weight=1)
        button_row1.grid_columnconfigure(2, weight=1)
        button_row1.grid_columnconfigure(3, weight=1)

        ttk.Button(button_row1, text="调试匹配",
                   command=self.debug_matching).grid(row=0, column=0, padx=5, sticky=tk.EW)
        ttk.Button(button_row1, text="预览重命名",
                   command=self.preview_excel_rename).grid(row=0, column=1, padx=5, sticky=tk.EW)
        ttk.Button(button_row1, text="执行重命名",
                   command=self.execute_excel_rename).grid(row=0, column=2, padx=5, sticky=tk.EW)
        ttk.Button(button_row1, text="清空预览",
                   command=self.clear_preview2).grid(row=0, column=3, padx=5, sticky=tk.EW)

        # 第二行按钮
        button_row2 = ttk.Frame(button_frame)
        button_row2.pack(fill=tk.X, pady=5)
        button_row2.grid_columnconfigure(0, weight=1)
        button_row2.grid_columnconfigure(1, weight=1)

        ttk.Button(button_row2, text="打开目标文件夹",
                   command=self.open_target_folder).grid(row=0, column=0, padx=5, sticky=tk.EW)
        ttk.Button(button_row2, text="统计文件夹文件",
                   command=self.count_files).grid(row=0, column=1, padx=5, sticky=tk.EW)

        # 输出区域 - 增大高度
        output_frame = ttk.LabelFrame(parent, text="输出结果", padding="10")
        output_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        output_frame.grid_rowconfigure(0, weight=1)
        output_frame.grid_columnconfigure(0, weight=1)

        # 创建文本区域和滚动条 - 增大高度
        text_frame = ttk.Frame(output_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)

        self.preview_text2 = tk.Text(text_frame, height=20)  # 增大高度
        scrollbar = ttk.Scrollbar(text_frame, command=self.preview_text2.yview)
        self.preview_text2.configure(yscrollcommand=scrollbar.set)

        self.preview_text2.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

    # ============ 通用方法 ============

    def select_folder1(self):
        """选择模式1的文件夹"""
        folder = filedialog.askdirectory(title="选择包含图片的文件夹")
        if folder:
            self.folder_path1.set(folder)

    def select_folder(self, var):
        """选择文件夹"""
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            var.set(folder)

    def select_excel(self):
        """选择Excel文件"""
        filetypes = [("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        filepath = filedialog.askopenfilename(title="选择Excel映射文件", filetypes=filetypes)
        if filepath:
            self.excel_path.set(filepath)

    def open_folder(self, folder_path):
        """打开文件夹"""
        if folder_path and os.path.exists(folder_path):
            if sys.platform == "darwin":  # macOS
                subprocess.run(["open", folder_path])
            elif sys.platform == "win32":  # Windows
                subprocess.run(["explorer", folder_path])
            else:  # Linux
                subprocess.run(["xdg-open", folder_path])
        else:
            messagebox.showwarning("提示", "请先选择有效的文件夹")

    def clean_filename(self, filename):
        """清理文件名：去除扩展名和空格"""
        if not filename:
            return ""

        # 去除空格
        filename = str(filename).strip()

        # 去除扩展名
        if '.' in filename:
            # 只保留最后一个点之前的部分
            filename = filename.rsplit('.', 1)[0]

        return filename

    def open_target_folder(self):
        """打开目标文件夹"""
        target_folder = self.target_folder_path.get()
        self.open_folder(target_folder)

    def count_files(self):
        """统计文件夹中的文件"""
        target_folder = self.target_folder_path.get()

        if not target_folder or not os.path.exists(target_folder):
            messagebox.showwarning("提示", "请先选择有效的目标文件夹")
            return

        target_path = Path(target_folder)
        files = list(target_path.glob("*.*"))

        self.preview_text2.delete(1.0, tk.END)
        self.preview_text2.insert(tk.END, f"=== 文件夹文件统计 ===\n")
        self.preview_text2.insert(tk.END, f"文件夹: {target_folder}\n")
        self.preview_text2.insert(tk.END, f"文件总数: {len(files)} 个\n")
        self.preview_text2.insert(tk.END, "="*60 + "\n")

        # 按扩展名统计
        ext_counts = {}
        for file in files:
            ext = file.suffix.lower()
            if ext in ext_counts:
                ext_counts[ext] += 1
            else:
                ext_counts[ext] = 1

        self.preview_text2.insert(tk.END, "按扩展名统计:\n")
        for ext, count in sorted(ext_counts.items()):
            self.preview_text2.insert(tk.END, f"  {ext}: {count} 个\n")

        # 显示前20个文件
        self.preview_text2.insert(tk.END, "\n前20个文件:\n")
        for i, file in enumerate(files[:20], 1):
            self.preview_text2.insert(tk.END, f"{i:2d}. {file.name}\n")

        if len(files) > 20:
            self.preview_text2.insert(tk.END, f"... 还有 {len(files)-20} 个文件未显示\n")

    # ============ 模式1方法 ============

    def preview_removal(self):
        """预览模式1的重命名"""
        folder = self.folder_path1.get()
        remove_str = self.remove_string.get()

        if not folder or not remove_str:
            messagebox.showerror("错误", "请先选择文件夹并输入要移除的字符串")
            return

        try:
            self.preview_text1.delete(1.0, tk.END)
            folder_path = Path(folder)
            files = list(folder_path.glob(f"*{remove_str}*"))

            if not files:
                messagebox.showwarning("警告", f"在文件夹中未找到包含 '{remove_str}' 的文件")
                return

            self.preview_text1.insert(tk.END, f"找到 {len(files)} 个文件，预览重命名：\n")
            self.preview_text1.insert(tk.END, "="*60 + "\n")

            for file in files[:20]:
                old_name = file.name
                new_name = old_name.replace(remove_str, '')
                self.preview_text1.insert(tk.END, f"{old_name}\n  → {new_name}\n")

            if len(files) > 20:
                self.preview_text1.insert(tk.END, f"\n... 还有 {len(files)-20} 个文件未显示\n")

            self.preview_text1.insert(tk.END, "="*60 + "\n")
            self.preview_text1.insert(tk.END, f"总计: {len(files)} 个文件将被重命名\n")

        except Exception as e:
            messagebox.showerror("错误", f"预览时发生错误：{str(e)}")

    def execute_removal(self):
        """执行模式1的重命名"""
        folder = self.folder_path1.get()
        remove_str = self.remove_string.get()

        if not folder or not remove_str:
            messagebox.showerror("错误", "请先选择文件夹并输入要移除的字符串")
            return

        try:
            folder_path = Path(folder)
            files = list(folder_path.glob(f"*{remove_str}*"))

            if not files:
                messagebox.showwarning("警告", f"在文件夹中未找到包含 '{remove_str}' 的文件")
                return

            # 确认对话框
            if not messagebox.askyesno("确认", f"确认要将 {len(files)} 个文件中的 '{remove_str}' 移除吗？"):
                return

            success_count = 0
            for file in files:
                try:
                    old_name = file.name
                    new_name = old_name.replace(remove_str, '')
                    new_path = file.parent / new_name

                    # 如果新文件名已存在，添加后缀避免冲突
                    counter = 1
                    while new_path.exists():
                        name_parts = new_name.rsplit('.', 1)
                        if len(name_parts) == 2:
                            new_name = f"{name_parts[0]}_{counter}.{name_parts[1]}"
                        else:
                            new_name = f"{new_name}_{counter}"
                        new_path = file.parent / new_name
                        counter += 1

                    file.rename(new_path)
                    success_count += 1
                except Exception as e:
                    self.preview_text1.insert(tk.END, f"重命名失败 {file.name}: {str(e)}\n")

            messagebox.showinfo("完成", f"重命名完成！成功: {success_count}/{len(files)} 个文件")

        except Exception as e:
            messagebox.showerror("错误", f"重命名时发生错误：{str(e)}")
            self.preview_text1.insert(tk.END, f"\n错误详情：{traceback.format_exc()}\n")

    def clear_preview1(self):
        """清空模式1的预览"""
        self.preview_text1.delete(1.0, tk.END)

    # ============ 模式2方法 ============

    def preview_excel_columns(self):
        """预览Excel文件的列名"""
        excel_path = self.excel_path.get()

        if not excel_path:
            messagebox.showwarning("提示", "请先选择Excel文件")
            return

        try:
            # 读取Excel文件
            df = pd.read_excel(excel_path)

            # 显示列名
            columns = df.columns.tolist()
            self.preview_text2.delete(1.0, tk.END)
            self.preview_text2.insert(tk.END, f"Excel文件: {os.path.basename(excel_path)}\n")
            self.preview_text2.insert(tk.END, "="*70 + "\n")
            self.preview_text2.insert(tk.END, f"共有 {len(df)} 行，{len(columns)} 列\n")
            self.preview_text2.insert(tk.END, "="*70 + "\n")
            self.preview_text2.insert(tk.END, "可用列名:\n")

            for i, col in enumerate(columns, 1):
                self.preview_text2.insert(tk.END, f"{i}. '{col}'\n")

            # 显示前几行数据
            self.preview_text2.insert(tk.END, "\n" + "="*70 + "\n")
            self.preview_text2.insert(tk.END, "前10行数据预览:\n")
            self.preview_text2.insert(tk.END, "="*70 + "\n")

            src_col = self.src_col_name.get()
            dst_col = self.dst_col_name.get()

            for idx, row in df.head(10).iterrows():
                self.preview_text2.insert(tk.END, f"\n第{idx+1}行:\n")
                if src_col in df.columns:
                    src_val = str(row[src_col])
                    clean_src = self.clean_filename(src_val)
                    self.preview_text2.insert(tk.END, f"  源列('{src_col}'): '{src_val}' → 清理后: '{clean_src}'\n")
                if dst_col in df.columns:
                    dst_val = str(row[dst_col])
                    clean_dst = self.clean_filename(dst_val)
                    self.preview_text2.insert(tk.END, f"  目标列('{dst_col}'): '{dst_val}' → 清理后: '{clean_dst}'\n")

        except Exception as e:
            messagebox.showerror("错误", f"读取Excel文件失败：{str(e)}")

    def test_mapping(self):
        """测试Excel映射关系"""
        excel_path = self.excel_path.get()

        if not excel_path:
            messagebox.showwarning("提示", "请先选择Excel文件")
            return

        try:
            # 读取Excel文件
            df = pd.read_excel(excel_path)
            src_col = self.src_col_name.get()
            dst_col = self.dst_col_name.get()

            if src_col not in df.columns:
                messagebox.showerror("错误",
                                     f"Excel文件中找不到列 '{src_col}'。可用列：{', '.join(df.columns.tolist())}")
                return

            if dst_col not in df.columns:
                messagebox.showerror("错误",
                                     f"Excel文件中找不到列 '{dst_col}'。可用列：{', '.join(df.columns.tolist())}")
                return

            self.preview_text2.delete(1.0, tk.END)
            self.preview_text2.insert(tk.END, f"=== 测试Excel映射关系 ===\n")
            self.preview_text2.insert(tk.END, f"源列: '{src_col}', 目标列: '{dst_col}'\n")
            self.preview_text2.insert(tk.END, "="*70 + "\n")

            # 测试映射关系
            mapping_issues = []
            mapping_count = 0

            for idx, row in df.iterrows():
                src_val = str(row[src_col])
                dst_val = str(row[dst_col])

                # 清理值
                clean_src = self.clean_filename(src_val)
                clean_dst = self.clean_filename(dst_val)

                self.preview_text2.insert(tk.END, f"\n第{idx+1}行:\n")
                self.preview_text2.insert(tk.END, f"  原始: '{src_val}' → '{dst_val}'\n")
                self.preview_text2.insert(tk.END, f"  清理后: '{clean_src}' → '{clean_dst}'\n")

                # 检查问题
                if not clean_src:
                    mapping_issues.append(f"第{idx+1}行: 源文件名为空")
                if not clean_dst:
                    mapping_issues.append(f"第{idx+1}行: 目标文件名为空")
                else:
                    mapping_count += 1

            self.preview_text2.insert(tk.END, "\n" + "="*70 + "\n")

            if mapping_issues:
                self.preview_text2.insert(tk.END, f"⚠️ 发现 {len(mapping_issues)} 个问题:\n")
                for issue in mapping_issues:
                    self.preview_text2.insert(tk.END, f"  - {issue}\n")
            else:
                self.preview_text2.insert(tk.END, f"✅ 映射关系测试通过！\n")

            self.preview_text2.insert(tk.END, f"有效映射记录数: {mapping_count}/{len(df)}\n")

        except Exception as e:
            messagebox.showerror("错误", f"测试映射关系失败：{str(e)}")

    def preview_full_excel(self):
        """查看Excel完整数据"""
        excel_path = self.excel_path.get()

        if not excel_path:
            messagebox.showwarning("提示", "请先选择Excel文件")
            return

        try:
            # 读取Excel文件
            df = pd.read_excel(excel_path)

            # 创建新窗口显示完整数据
            excel_window = tk.Toplevel(self.root)
            excel_window.title(f"Excel数据预览 - {os.path.basename(excel_path)}")
            excel_window.geometry("900x600")

            # 配置网格权重
            excel_window.grid_rowconfigure(0, weight=1)
            excel_window.grid_columnconfigure(0, weight=1)

            # 添加Treeview显示数据
            tree_frame = ttk.Frame(excel_window, padding="10")
            tree_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            tree_frame.grid_rowconfigure(0, weight=1)
            tree_frame.grid_columnconfigure(0, weight=1)

            # 创建Treeview
            tree = ttk.Treeview(tree_frame, show="headings")

            # 定义列
            columns = df.columns.tolist()
            tree["columns"] = columns

            # 设置列标题
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=150, minwidth=50)

            # 添加滚动条
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

            tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
            hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))

            # 添加数据
            for idx, row in df.iterrows():
                values = [str(row[col]) for col in columns]
                tree.insert("", "end", values=values)

            # 设置网格权重
            tree_frame.grid_rowconfigure(0, weight=1)
            tree_frame.grid_columnconfigure(0, weight=1)

            # 关闭按钮
            close_button = ttk.Button(excel_window, text="关闭", command=excel_window.destroy)
            close_button.grid(row=1, column=0, pady=10)

        except Exception as e:
            messagebox.showerror("错误", f"读取Excel文件失败：{str(e)}")

    def get_mapping_dict(self):
        """获取清理后的映射字典"""
        excel_path = self.excel_path.get()
        src_col = self.src_col_name.get()
        dst_col = self.dst_col_name.get()

        if not excel_path:
            return None

        try:
            df = pd.read_excel(excel_path)

            if src_col not in df.columns or dst_col not in df.columns:
                return None

            mapping = {}
            for _, row in df.iterrows():
                src_val = str(row[src_col]).strip()
                dst_val = str(row[dst_col]).strip()

                # 清理文件名（去除扩展名）
                clean_src = self.clean_filename(src_val)
                clean_dst = self.clean_filename(dst_val)

                if clean_src and clean_dst:
                    mapping[clean_src] = clean_dst

            return mapping

        except Exception as e:
            messagebox.showerror("错误", f"读取Excel映射失败：{str(e)}")
            return None

    def debug_matching(self):
        """调试匹配过程"""
        target_folder = self.target_folder_path.get()
        excel_path = self.excel_path.get()

        if not all([target_folder, excel_path]):
            messagebox.showerror("错误", "请选择目标文件夹和Excel文件")
            return

        mapping = self.get_mapping_dict()
        if not mapping:
            return

        target_path = Path(target_folder)

        self.preview_text2.delete(1.0, tk.END)
        self.preview_text2.insert(tk.END, f"=== 调试匹配过程 ===\n")
        self.preview_text2.insert(tk.END, f"时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        self.preview_text2.insert(tk.END, f"匹配模式: {self.match_mode.get()}\n")
        self.preview_text2.insert(tk.END, "="*80 + "\n")

        # 显示映射表
        self.preview_text2.insert(tk.END, f"Excel映射表 ({len(mapping)} 条):\n")
        self.preview_text2.insert(tk.END, "-"*40 + "\n")

        for i, (src, dst) in enumerate(mapping.items(), 1):
            self.preview_text2.insert(tk.END, f"{i:2d}. '{src}' → '{dst}'\n")

        # 显示目标文件夹文件
        target_files = list(target_path.glob("*.*"))
        self.preview_text2.insert(tk.END, f"\n目标文件夹文件 ({len(target_files)} 个):\n")
        self.preview_text2.insert(tk.END, "-"*40 + "\n")

        for i, file in enumerate(target_files[:30], 1):  # 只显示前30个
            file_name = file.name
            file_stem = file.stem

            self.preview_text2.insert(tk.END, f"{i:2d}. '{file_name}' (主干: '{file_stem}')\n")

        if len(target_files) > 30:
            self.preview_text2.insert(tk.END, f"... 还有 {len(target_files)-30} 个文件未显示\n")

        # 开始匹配测试
        self.preview_text2.insert(tk.END, "\n" + "="*80 + "\n")
        self.preview_text2.insert(tk.END, "匹配测试结果:\n")
        self.preview_text2.insert(tk.END, "-"*40 + "\n")

        matched_count = 0
        match_mode = self.match_mode.get()

        for file in target_files[:30]:  # 只测试前30个
            file_name = file.name
            file_stem = file.stem
            file_ext = file.suffix

            self.preview_text2.insert(tk.END, f"\n文件: '{file_name}'\n")
            self.preview_text2.insert(tk.END, f"  主干: '{file_stem}', 扩展名: '{file_ext}'\n")

            found_match = False

            for excel_src, excel_dst in mapping.items():
                # 检查匹配
                if match_mode == "exact":
                    match = (file_stem == excel_src)
                elif match_mode == "contains":
                    match = (excel_src in file_stem)
                else:
                    match = False

                if match:
                    # 确定新文件名
                    if self.keep_original_extension.get():
                        new_name = excel_dst + file_ext
                    else:
                        new_name = excel_dst + ".jpg"

                    self.preview_text2.insert(tk.END, f"  ✅ 匹配成功！\n")
                    self.preview_text2.insert(tk.END, f"      Excel源: '{excel_src}'\n")
                    self.preview_text2.insert(tk.END, f"      Excel目标: '{excel_dst}'\n")
                    self.preview_text2.insert(tk.END, f"      新文件名: '{new_name}'\n")

                    matched_count += 1
                    found_match = True
                    break

            if not found_match:
                self.preview_text2.insert(tk.END, f"  ❌ 未找到任何匹配\n")

        self.preview_text2.insert(tk.END, "\n" + "="*80 + "\n")
        total_tested = min(30, len(target_files))
        self.preview_text2.insert(tk.END, f"匹配统计: 测试 {total_tested} 个文件，找到 {matched_count} 个匹配\n")

        if matched_count == 0:
            self.preview_text2.insert(tk.END, f"\n建议:\n")
            self.preview_text2.insert(tk.END, f"1. 检查Excel中的源文件名是否正确\n")
            self.preview_text2.insert(tk.END, f"2. 尝试'包含匹配'模式\n")
            self.preview_text2.insert(tk.END, f"3. 检查文件名是否有隐藏字符\n")

    def preview_excel_rename(self):
        """预览模式2的重命名"""
        target_folder = self.target_folder_path.get()
        excel_path = self.excel_path.get()

        if not all([target_folder, excel_path]):
            messagebox.showerror("错误", "请选择目标文件夹和Excel文件")
            return

        mapping = self.get_mapping_dict()
        if not mapping:
            return

        target_path = Path(target_folder)

        self.preview_text2.delete(1.0, tk.END)
        self.preview_text2.insert(tk.END, f"=== 重命名预览 ===\n")
        self.preview_text2.insert(tk.END, f"匹配模式: {self.match_mode.get()}\n")
        self.preview_text2.insert(tk.END, f"Excel映射表中有 {len(mapping)} 条记录\n")
        self.preview_text2.insert(tk.END, f"目标文件夹: {target_folder}\n")
        self.preview_text2.insert(tk.END, "="*60 + "\n")

        # 预览重命名
        matched_count = 0
        target_files = list(target_path.glob("*.*"))
        match_mode = self.match_mode.get()

        for file in target_files[:30]:  # 只预览前30个
            file_name = file.name
            file_stem = file.stem
            file_ext = file.suffix

            # 尝试匹配
            matched_key = None
            for key in mapping.keys():
                if match_mode == "exact":
                    if file_stem == key:
                        matched_key = key
                        break
                elif match_mode == "contains":
                    if key in file_stem:
                        matched_key = key
                        break

            if matched_key:
                target_base = mapping[matched_key]

                # 确定新文件名
                if self.keep_original_extension.get():
                    new_name = target_base + file_ext
                else:
                    new_name = target_base + ".jpg"

                self.preview_text2.insert(tk.END, f"✅ {file_name}\n")
                self.preview_text2.insert(tk.END, f"   → {new_name}\n")
                self.preview_text2.insert(tk.END, f"   匹配方式: {match_mode} ('{matched_key}')\n\n")
                matched_count += 1
            else:
                self.preview_text2.insert(tk.END, f"❌ {file_name} - 未找到匹配的映射\n\n")

        if len(target_files) > 30:
            remaining = len(target_files) - 30
            self.preview_text2.insert(tk.END, f"... 还有 {remaining} 个文件未显示\n")

        self.preview_text2.insert(tk.END, "="*60 + "\n")
        total_previewed = min(30, len(target_files))
        self.preview_text2.insert(tk.END,
                                  f"预览统计：{matched_count}/{total_previewed} 个文件找到匹配\n")

        if matched_count > 0:
            self.preview_text2.insert(tk.END, f"\n✅ 可以执行重命名！\n")
        else:
            self.preview_text2.insert(tk.END, f"\n⚠️ 未找到匹配，请尝试以下方法：\n")
            self.preview_text2.insert(tk.END, f"1. 点击'调试匹配'查看详细匹配过程\n")
            self.preview_text2.insert(tk.END, f"2. 调整匹配模式\n")
            self.preview_text2.insert(tk.END, f"3. 检查Excel列名是否正确\n")

    def execute_excel_rename(self):
        """执行模式2的重命名"""
        target_folder = self.target_folder_path.get()
        excel_path = self.excel_path.get()

        if not all([target_folder, excel_path]):
            messagebox.showerror("错误", "请选择目标文件夹和Excel文件")
            return

        mapping = self.get_mapping_dict()
        if not mapping:
            return

        target_path = Path(target_folder)
        target_files = list(target_path.glob("*.*"))

        # 确认对话框
        confirm_msg = (
            f"确认要开始批量重命名吗？\n"
            f"Excel映射表中有 {len(mapping)} 条记录\n"
            f"目标文件夹有 {len(target_files)} 个文件\n"
            f"匹配模式: {self.match_mode.get()}"
        )

        if not messagebox.askyesno("确认执行", confirm_msg):
            return

        success_count = 0
        error_count = 0
        no_match_count = 0

        self.preview_text2.delete(1.0, tk.END)
        self.preview_text2.insert(tk.END, "=== 开始批量重命名 ===\n")
        self.preview_text2.insert(tk.END, f"时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        self.preview_text2.insert(tk.END, "="*60 + "\n")

        # 创建备份文件夹
        backup_dir = target_path / "backup_rename"
        backup_dir.mkdir(exist_ok=True)
        self.preview_text2.insert(tk.END, f"备份文件夹: {backup_dir}\n\n")

        match_mode = self.match_mode.get()

        for file in target_files:
            try:
                file_name = file.name
                file_stem = file.stem
                file_ext = file.suffix

                # 尝试匹配
                matched_key = None
                for key in mapping.keys():
                    if match_mode == "exact":
                        if file_stem == key:
                            matched_key = key
                            break
                    elif match_mode == "contains":
                        if key in file_stem:
                            matched_key = key
                            break

                if not matched_key:
                    no_match_count += 1
                    continue

                target_base = mapping[matched_key]

                # 确定新文件名
                if self.keep_original_extension.get():
                    new_name = target_base + file_ext
                else:
                    new_name = target_base + ".jpg"

                old_path = file
                new_path = file.parent / new_name

                # 如果新文件就是原文件，跳过
                if new_path == old_path:
                    self.preview_text2.insert(tk.END, f"⏭️ {file_name} → 已是目标名称，跳过\n")
                    continue

                # 备份原文件
                backup_path = backup_dir / file_name
                import shutil
                shutil.copy2(file, backup_path)

                # 如果新文件名已存在，添加后缀避免冲突
                counter = 1
                while new_path.exists():
                    name_parts = new_name.rsplit('.', 1)
                    if len(name_parts) == 2:
                        new_name = f"{name_parts[0]}_{counter}.{name_parts[1]}"
                    else:
                        new_name = f"{new_name}_{counter}"
                    new_path = file.parent / new_name
                    counter += 1

                # 执行重命名
                file.rename(new_path)

                self.preview_text2.insert(tk.END, f"✅ {file_name} → {new_name}\n")
                success_count += 1

            except Exception as e:
                self.preview_text2.insert(tk.END, f"❌ {file_name} → 错误: {str(e)}\n")
                error_count += 1

        self.preview_text2.insert(tk.END, "\n" + "="*60 + "\n")
        self.preview_text2.insert(tk.END, f"重命名完成！\n")
        self.preview_text2.insert(tk.END, f"成功: {success_count} 个文件\n")
        self.preview_text2.insert(tk.END, f"失败: {error_count} 个文件\n")
        self.preview_text2.insert(tk.END, f"无匹配: {no_match_count} 个文件\n")
        self.preview_text2.insert(tk.END, f"总计: {len(target_files)} 个文件\n")

        if success_count > 0:
            self.preview_text2.insert(tk.END, f"\n✅ 备份文件保存在: {backup_dir}\n")

        # 保存日志文件
        log_file = target_path / f"rename_log_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.txt"
        log_content = self.preview_text2.get(1.0, tk.END)
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(log_content)

        self.preview_text2.insert(tk.END, f"\n📝 详细日志保存到: {log_file}\n")

        messagebox.showinfo("完成",
                            f"重命名完成！\n\n"
                            f"成功: {success_count} 个文件\n"
                            f"失败: {error_count} 个文件\n"
                            f"无匹配: {no_match_count} 个文件\n\n"
                            f"备份文件在: {backup_dir}\n"
                            f"详细日志: {log_file}")

    def clear_preview2(self):
        """清空模式2的预览"""
        self.preview_text2.delete(1.0, tk.END)

    def run(self):
        """运行主程序"""
        self.root.mainloop()

def main():
    """主函数"""
    try:
        app = BatchRenamer()
        app.run()
    except Exception as e:
        messagebox.showerror("程序错误", f"程序启动失败：{str(e)}\n{traceback.format_exc()}")

if __name__ == "__main__":
    main()