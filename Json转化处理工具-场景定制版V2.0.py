import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import json
import re
from datetime import datetime
import os
import threading
from statistics import mode, StatisticsError
import numpy as np
import sys

# Custom JSON encoder to handle NumPy types
class NumpyEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (np.int_, np.intc, np.intp, np.int8,
                            np.int16, np.int32, np.int64, np.uint8,
                            np.uint16, np.uint32, np.uint64)):
            return int(obj)
        elif isinstance(obj, (np.float_, np.float16, np.float32, np.float64)):
            return float(obj)
        elif isinstance(obj, np.bool_):
            return bool(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        return super().default(obj)

class ExcelJSONProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("虚假类警告信Json转换处理脚本")
        self.root.geometry("850x800")
        self.root.minsize(800, 750)
        self.root.configure(bg='#f8f8f8')

        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', background='#f0f0f0')
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabelFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0')
        style.configure('Title.TLabel', font=('Microsoft YaHei UI', 16, 'bold'), background='#f0f0f0')
        style.configure('Heading.TLabel', font=('Microsoft YaHei UI', 10, 'bold'), background='#f0f0f0')
        style.configure('Accent.TButton', font=('Microsoft YaHei UI', 10, 'bold'))

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Title ---
        title_label = ttk.Label(main_frame, text="虚假类警告信Json转换处理脚本", style='Title.TLabel', anchor='center')
        title_label.pack(pady=5)
        subtitle_label = ttk.Label(main_frame, text="False Warning Letter Json Conversion & Processing Script", anchor='center')
        subtitle_label.pack(pady=(0, 15))

        # --- File Selection ---
        file_frame = ttk.LabelFrame(main_frame, text="📁 文件选择 / File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        self.file_path_var = tk.StringVar()
        file_entry_label = ttk.Label(file_frame, text="源Excel文件 / Source Excel File:")
        file_entry_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, state='readonly', width=80)
        file_entry.grid(row=1, column=0, sticky=tk.EW, padx=5)
        select_button = ttk.Button(file_frame, text="浏览 / Browse...", command=self.select_file)
        select_button.grid(row=1, column=1, padx=5)
        file_frame.columnconfigure(0, weight=1)

        # --- Output Configuration ---
        config_frame = ttk.LabelFrame(main_frame, text="⚙️ 输出配置 / Output Configuration", padding="10")
        config_frame.pack(fill=tk.X, pady=5)

        # Output Directory Selection
        output_dir_label = ttk.Label(config_frame, text="输出目录 / Output Directory:", style='Heading.TLabel')
        output_dir_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.output_dir_var = tk.StringVar()
        output_dir_entry = ttk.Entry(config_frame, textvariable=self.output_dir_var, state='readonly', width=60)
        output_dir_entry.grid(row=1, column=0, sticky=tk.EW, padx=5, columnspan=2)
        select_dir_button = ttk.Button(config_frame, text="选择目录 / Select Dir...", command=self.select_output_directory)
        select_dir_button.grid(row=1, column=2, padx=5)
        config_frame.columnconfigure(0, weight=1)

        # Prefix
        prefix_label = ttk.Label(config_frame, text="文件名优选前缀 / Preferred Filename Prefix:", style='Heading.TLabel')
        prefix_label.grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.prefix_var = tk.StringVar(value="虚假妥投警告信")
        self.prefix_combo = ttk.Combobox(config_frame, textvariable=self.prefix_var, width=40)
        self.prefix_combo['values'] = ("虚假妥投警告信", "虚假标记警告信", "False Delivery Warning List", "False Marking Warning List")
        self.prefix_combo.grid(row=3, column=0, sticky=tk.W, padx=5)

        # Suffix
        suffix_label = ttk.Label(config_frame, text="文件名后缀 / Filename Suffix:", style='Heading.TLabel')
        suffix_label.grid(row=2, column=1, sticky=tk.W, padx=20, pady=2)
        self.suffix_type_var = tk.StringVar(value="auto")
        auto_suffix_radio = ttk.Radiobutton(config_frame, text="自动提取日期 (MMDD) / Auto-extract Date (MMDD)", variable=self.suffix_type_var, value="auto")
        auto_suffix_radio.grid(row=3, column=1, sticky=tk.W, padx=20)
        custom_suffix_radio = ttk.Radiobutton(config_frame, text="自定义后缀 / Custom Suffix:", variable=self.suffix_type_var, value="custom")
        custom_suffix_radio.grid(row=4, column=1, sticky=tk.W, padx=20)
        self.custom_suffix_var = tk.StringVar()
        custom_suffix_entry = ttk.Entry(config_frame, textvariable=self.custom_suffix_var, width=20)
        custom_suffix_entry.grid(row=4, column=1, sticky=tk.W, padx=160)

        # --- Data Correction Config ---
        correction_frame = ttk.LabelFrame(main_frame, text="🔧 违规类型代码设置 / Violation Type Code Settings", padding="10")
        correction_frame.pack(fill=tk.X, pady=5)
        
        type_label = ttk.Label(correction_frame, text="违规类型代码 / Violation Type Code:", style='Heading.TLabel')
        type_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.violation_type_int_var = tk.StringVar(value="19")
        type_entry = ttk.Entry(correction_frame, textvariable=self.violation_type_int_var, width=15)
        type_entry.grid(row=1, column=0, sticky=tk.W, padx=5)

        # --- Actions and Progress ---
        action_frame = ttk.Frame(main_frame, padding="10")
        action_frame.pack(fill=tk.X, pady=10)
        self.process_button = ttk.Button(action_frame, text="🚀 开始处理数据 / Start Processing", command=self.start_processing, style='Accent.TButton')
        self.process_button.pack(side=tk.LEFT, padx=10)
        clear_button = ttk.Button(action_frame, text="🗑️ 清空配置 / Clear Fields", command=self.clear_fields)
        clear_button.pack(side=tk.LEFT, padx=10)
        
        progress_frame = ttk.LabelFrame(main_frame, text="📊 处理进度监控 / Processing Progress Monitor", padding="10")
        progress_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.status_var = tk.StringVar(value="就绪等待 / Ready")
        status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        status_label.pack(anchor=tk.W)
        self.log_text = tk.Text(progress_frame, height=15, wrap=tk.WORD, font=('Consolas', 9), bg='#f8f8f8')
        log_scrollbar = ttk.Scrollbar(progress_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, before=self.log_text)
        self.add_log("系统初始化完成 / System initialized.")

    def add_log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def select_file(self):
        file_path = filedialog.askopenfilename(title="选择Excel文件 / Select Excel File", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            self.file_path_var.set(file_path)
            self.add_log(f"已选择源文件 / Source file selected: {os.path.basename(file_path)}")

    def select_output_directory(self):
        dir_path = filedialog.askdirectory(title="选择保存目录 / Select Output Directory")
        if dir_path:
            self.output_dir_var.set(dir_path)
            self.add_log(f"已选择保存目录 / Output directory selected: {dir_path}")

    def clear_fields(self):
        self.file_path_var.set("")
        self.output_dir_var.set("")
        self.prefix_var.set("虚假妥投警告信")
        self.suffix_type_var.set("auto")
        self.custom_suffix_var.set("")
        self.violation_type_int_var.set("19")
        self.log_text.delete(1.0, tk.END)
        self.status_var.set("就绪等待 / Ready")
        self.add_log("所有字段已清空 / All fields cleared.")

    def normalize_date(self, date_value):
        if pd.isna(date_value): return None
        if isinstance(date_value, (pd.Timestamp, datetime)):
            return date_value.strftime('%Y-%m-%d')
        try:
            return pd.to_datetime(date_value).strftime('%Y-%m-%d')
        except (ValueError, TypeError):
            return None

    def extract_date_from_filename(self, filename):
        patterns = [r'(\d{1,2})月(\d{1,2})日', r'(\d{1,2})-(\d{1,2})', r'(\d{4})(\d{2})(\d{2})', r'(\d{2})(\d{2})']
        for pattern in patterns:
            match = re.search(pattern, filename)
            if match:
                groups = match.groups()
                if len(groups) == 3: # YYYYMMDD
                    return f"{int(groups[1]):02d}{int(groups[2]):02d}"
                if len(groups) == 2: # MMDD
                    return f"{int(groups[0]):02d}{int(groups[1]):02d}"
        return datetime.now().strftime("%m%d")

    def standardize_violation_type(self, value):
        if pd.isna(value):
            return value
        
        value_str = str(value).lower()
        
        if any(keyword in value_str for keyword in ['严厉', '严重', 'stern', 'severe']):
            return "严厉Stern"
        elif any(keyword in value_str for keyword in ['口述', '口头', 'verbal', 'oral']):
            return "口述Verbal"
        else:
            return value

    def determine_violation_type(self, series):
        # 基于预处理后的标准化值进行识别
        series_str = series.astype(str)
        if (series_str == "口述Verbal").all():
            return "口述Verbal"
        if (series_str == "严厉Stern").all():
            return "严厉Stern"
        return series.iloc[0]

    def determine_upgraded_violation_type(self, series):
        # 基于预处理后的标准化值进行识别和升级处理
        series_str = series.astype(str)
        if (series_str == "口述Verbal").all():
            return "严厉Stern"  # "升级处罚"规则
        if (series_str == "严厉Stern").all():
            return "严厉Stern"
        return series.iloc[0]

    def preprocess_data(self, details_df, auxiliary_df, filename):
        self.add_log("开始数据预处理 / Starting data preprocessing...")
        
        # 0. 初始清理
        details_df['Violation date'] = pd.to_datetime(details_df['Violation date'], errors='coerce').dt.date

        # 1. 跨Sheet去重
        if not auxiliary_df.empty and 'false_bill_num' in auxiliary_df.columns and 'false_bill_num' in details_df.columns:
            initial_count = len(details_df)
            duplicates = details_df['false_bill_num'].isin(auxiliary_df['false_bill_num'])
            details_df = details_df[~duplicates].copy()
            self.add_log(f"步骤1 (跨表去重): 移除了 {initial_count - len(details_df)} 条记录 / Step 1 (Cross-sheet dedup): Removed {initial_count - len(details_df)} records.")

        # 2. Violation type统一化
        details_df['Violation type'] = details_df['Violation type'].apply(self.standardize_violation_type)
        self.add_log("步骤2: Violation type 值已标准化 / Step 2: Violation type values standardized.")

        # 3. 执行完全合并
        self.add_log("步骤3: 执行完全合并 / Step 3: Performing Exact Merge...")
        processed_dfs_pass1 = []
        exact_merged_indices = []  # 记录参与完全合并的索引

        groups_to_process_pass1 = details_df.groupby(['Employee ID', 'Violation date', 'Violation type'])
        
        for group_keys, group in groups_to_process_pass1:
            if len(group) > 1:
                # 记录参与完全合并的所有索引
                exact_merged_indices.extend(group.index.tolist())
                new_row = group.iloc[0].copy()
                new_row['false_bill_num'] = ",".join(group['false_bill_num'].astype(str))
                new_row['false_num'] = group['false_num'].sum()
                new_row['Violation type'] = self.determine_violation_type(group['Violation type'])
                processed_dfs_pass1.append(new_row.to_frame().T)
        
        # 获取未参与完全合并的记录
        unmerged_after_exact = details_df.loc[~details_df.index.isin(exact_merged_indices)]
        
        # 完全合并后的结果
        exact_merge_results = pd.concat(processed_dfs_pass1, ignore_index=True) if processed_dfs_pass1 else pd.DataFrame()
        
        self.add_log(f"完全合并: 合并了 {len(exact_merged_indices)} 条记录为 {len(exact_merge_results)} 条，剩余 {len(unmerged_after_exact)} 条未合并记录 / Exact Merge: Merged {len(exact_merged_indices)} records into {len(exact_merge_results)} records, {len(unmerged_after_exact)} records remain unmerged.")

        # 4. 执行部分合并（只对未参与完全合并的记录进行处理）
        self.add_log("步骤4: 执行部分合并 / Step 4: Performing Partial Merge...")
        processed_dfs_pass2 = []
        partial_merged_indices = []  # 记录参与部分合并的索引

        groups_to_process_pass2 = unmerged_after_exact.groupby(['Employee ID', 'Violation type'])

        for group_keys, group in groups_to_process_pass2:
            if len(group) >= 3:
                self.add_log(f"部分合并: 正在处理 Employee ID {group_keys[0]} 的 {len(group)} 条记录 / Partial Merge: Processing {len(group)} records for Employee ID {group_keys[0]}.")
                # 记录参与部分合并的所有索引
                partial_merged_indices.extend(group.index.tolist())

                # --- 调整部分开始 ---
                # 创建一个包含单号和日期的列表，用于拼接
                false_bill_nums_with_date = [
                    f"{row['false_bill_num']}({self.normalize_date(row['Violation date'])})"
                    if pd.notna(row['Violation date']) else str(row['false_bill_num'])
                    for _, row in group.iterrows()
                ]

                new_row = group.iloc[0].copy()
                new_row['false_bill_num'] = ",".join(false_bill_nums_with_date)
                # --- 调整部分结束 ---

                new_row['false_num'] = group['false_num'].sum()
                new_row['Violation type'] = self.determine_upgraded_violation_type(group['Violation type'])
                
                month_day_from_filename = self.extract_date_from_filename(filename)
                valid_dates = group['Violation date'].dropna()
                try:
                    year_mode = mode(d.year for d in valid_dates)
                except StatisticsError:
                    year_mode = datetime.now().year
                new_row['Violation date'] = pd.to_datetime(f"{year_mode}-{month_day_from_filename[:2]}-{month_day_from_filename[2:]}").date()
                
                processed_dfs_pass2.append(new_row.to_frame().T)

        # 获取既未参与完全合并也未参与部分合并的记录
        unmerged_final = unmerged_after_exact.loc[~unmerged_after_exact.index.isin(partial_merged_indices)]
        
        # 部分合并后的结果
        partial_merge_results = pd.concat(processed_dfs_pass2, ignore_index=True) if processed_dfs_pass2 else pd.DataFrame()
        
        self.add_log(f"部分合并: 合并了 {len(partial_merged_indices)} 条记录为 {len(partial_merge_results)} 条，最终剩余 {len(unmerged_final)} 条未合并记录 / Partial Merge: Merged {len(partial_merged_indices)} records into {len(partial_merge_results)} records, {len(unmerged_final)} records remain unmerged.")

        # 合并最终结果
        final_parts = []
        if not exact_merge_results.empty:
            final_parts.append(exact_merge_results)
        if not partial_merge_results.empty:
            final_parts.append(partial_merge_results)
        if not unmerged_final.empty:
            final_parts.append(unmerged_final)
        
        final_df = pd.concat(final_parts, ignore_index=True) if final_parts else pd.DataFrame()
        self.add_log(f"数据预处理完成，最终剩余 {len(final_df)} 条记录 / Preprocessing finished, {len(final_df)} records remaining.")
        return final_df

    def create_json_column(self, df):
        self.add_log("正在创建JSON列 / Creating JSON column...")
        cols_to_json = [col for col in df.columns if col not in ['Employee ID', 'Violation date', 'Violation type', 'Violation details']]
        
        def to_json(row):
            # Use NumpyEncoder to handle non-standard types like int64
            data = {col: row[col] for col in cols_to_json if pd.notna(row[col])}
            if data:
                return json.dumps(data, ensure_ascii=False, cls=NumpyEncoder)
            return None

        df['Violation details'] = df.apply(to_json, axis=1)
        return df[['Employee ID', 'Violation date', 'Violation type', 'Violation details']]

    def parse_json_details(self, df):
        """
        Parse JSON values in 'Violation details' column and create separate columns
        for false_type, false_num, and false_bill_num
        """
        self.add_log("正在解析JSON详情字段 / Parsing JSON details field...")
        
        # Initialize new columns
        df['false_type'] = None
        df['false_num'] = None
        df['false_bill_num'] = None
        
        # Parse each row's JSON data
        for idx, row in df.iterrows():
            if pd.notna(row['Violation details']):
                try:
                    json_data = json.loads(row['Violation details'])
                    df.at[idx, 'false_type'] = json_data.get('false_type', None)
                    df.at[idx, 'false_num'] = json_data.get('false_num', None)
                    df.at[idx, 'false_bill_num'] = json_data.get('false_bill_num', None)
                except (json.JSONDecodeError, TypeError) as e:
                    self.add_log(f"警告: 行 {idx} 的JSON解析失败 / Warning: Failed to parse JSON at row {idx}: {e}")
        
        # Remove the original Violation details column
        df = df.drop('Violation details', axis=1)
        
        # Reorder columns to put the new fields after the main fields
        cols = ['Employee ID', 'Violation date', 'Violation type', 'false_type', 'false_num', 'false_bill_num']
        df = df[cols]
        
        self.add_log("JSON详情解析完成 / JSON details parsing completed.")
        return df

    def correct_data(self, df, violation_type_int):
        self.add_log("正在进行最终数据纠正 / Performing final data correction...")
        
        # Create original_type_df with parsed JSON fields
        original_type_df = df.copy()
        original_type_df = self.parse_json_details(original_type_df)
        
        # Process main df
        df['Violation date'] = df['Violation date'].apply(self.normalize_date)
        df['Violation type'] = violation_type_int
        self.add_log(f"Violation type 已统一为 {violation_type_int} / Violation type standardized to {violation_type_int}.")
        
        return df, original_type_df

    def start_processing(self):
        if not self.file_path_var.get():
            messagebox.showerror("错误 / Error", "请先选择一个Excel文件 / Please select an Excel file first.")
            return
        
        if not self.output_dir_var.get():
            messagebox.showerror("错误 / Error", "请先选择一个输出目录 / Please select an output directory first.")
            return
        
        try:
            int(self.violation_type_int_var.get())
        except ValueError:
            messagebox.showerror("错误 / Error", "Violation Type整数值必须为有效数字 / The integer value for Violation Type must be a valid number.")
            return

        self.process_button.config(state="disabled", text="处理中 / Processing...")
        thread = threading.Thread(target=self.process_excel, daemon=True)
        thread.start()

    def process_excel(self):
        try:
            try:
                import xlsxwriter
            except ImportError:
                raise ImportError("`xlsxwriter` module is not installed. Please install it with `pip install xlsxwriter`.")

            input_file = self.file_path_var.get()
            output_dir = self.output_dir_var.get()
            self.status_var.set("正在读取文件 / Reading file...")
            
            xls = pd.ExcelFile(input_file)
            if 'details' not in xls.sheet_names:
                raise ValueError("Excel文件中必须包含'details'工作表 / Excel file must contain a 'details' sheet.")
            
            details_df = pd.read_excel(xls, sheet_name='details')
            auxiliary_df = pd.read_excel(xls, sheet_name='auxiliary') if 'auxiliary' in xls.sheet_names else pd.DataFrame()
            self.add_log(f"读取到 {len(details_df)} 条 'details' 记录和 {len(auxiliary_df)} 条 'auxiliary' 记录 / Read {len(details_df)} 'details' records and {len(auxiliary_df)} 'auxiliary' records.")

            # 预处理
            processed_df = self.preprocess_data(details_df, auxiliary_df, os.path.basename(input_file))

            # JSON转换
            json_df = self.create_json_column(processed_df)

            # 数据纠正
            final_df, original_type_df = self.correct_data(json_df, int(self.violation_type_int_var.get()))

            # 生成文件名并保存
            self.status_var.set("正在生成并保存文件 / Generating and saving file...")
            prefix = self.prefix_var.get()
            if self.suffix_type_var.get() == "custom":
                suffix = self.custom_suffix_var.get() if self.custom_suffix_var.get() else datetime.now().strftime("%m%d")
            else:
                suffix = self.extract_date_from_filename(os.path.basename(input_file))
            
            output_filename = f"{prefix}-{suffix}.xlsx"
            output_path = os.path.join(output_dir, output_filename)

            with pd.ExcelWriter(output_path, engine='xlsxwriter', date_format='YYYY-MM-DD', datetime_format='YYYY-MM-DD') as writer:
                final_df.to_excel(writer, sheet_name='details', index=False)
                original_type_df.to_excel(writer, sheet_name='details_original_type', index=False)
                
                workbook = writer.book
                worksheet_to_hide = writer.sheets['details_original_type']
                worksheet_to_hide.hide()

            self.add_log(f"处理完成！文件已保存至 / Processing complete! File saved to: {output_path}")
            self.status_var.set("处理完成！ / Processing complete!")
            messagebox.showinfo("成功 / Success", f"文件已成功处理并保存！\nFile processed and saved successfully!\n\n路径 / Path: {output_path}")

        except ImportError as e:
            self.add_log(f"模块缺失错误 / Missing module error: {e}")
            self.add_log("请安装 xlsxwriter 库来解决此问题： pip install xlsxwriter")
            messagebox.showerror("错误 / Error", f"处理过程中发生错误：\nAn error occurred during processing:\n\n{e}\n\n请安装 xlsxwriter 库：\n Please install the xlsxwriter library: pip install xlsxwriter")
        except Exception as e:
            import traceback
            self.add_log(f"发生错误 / An error occurred: {e}")
            self.add_log(f"Traceback: {traceback.format_exc()}")
            self.status_var.set(f"错误 / Error: {e}")
            messagebox.showerror("错误 / Error", f"处理过程中发生错误：\nAn error occurred during processing:\n\n{e}")
        finally:
            self.process_button.config(state="normal", text="🚀 开始处理数据 / Start Processing")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelJSONProcessor()
    app.run()
