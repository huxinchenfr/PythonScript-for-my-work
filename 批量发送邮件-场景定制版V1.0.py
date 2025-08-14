import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import smtplib
import threading
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email import encoders
import os
import re

class EmailSenderApp(tk.Tk):
    """
    自动化邮件批量发送工具-历史警告信数据处理专版
    """
    def __init__(self):
        super().__init__()
        self.title("自动化邮件批量发送工具-历史警告信数据处理专版")
        self.geometry("1100x900")

        # --- Internal representation of column names ---
        self.COLUMN_MAP = {}
        self.KEY_COLS = {
            'id': ('staff id', '工号'),
            'name': ('staff name', '姓名'),
            'warning_type': ('warning type', 'warning letter', '警告类型'),
            'area': ('area', '大区'),
            'district': ('district', '片区'),
            'branch': ('branch', '网点'),
            'ops': ('ops', '部门'),
            'position': ('position', '职位'),
            'status': ('work status', '在职状态'),
            'employment_type': ('employment', '雇佣类型'),
            'sending_status': ('sending status', '发送状态')
        }
        self.CRITICAL_COLS = ['id', 'warning_type']
        
        # --- English processing configuration ---
        self.english_processing_values = set()  # 用户配置的需要英文处理的字段值
        self.split_field_values = []  # 拆分字段的所有可用值

        # --- Sheet name mappings (B. Sheet命名调整) ---
        self.SHEET_NAMES = {
            'chinese': {
                "details": "details",
                "2x_stern": "累计2次严厉警告员工明细",
                "3x_stern": "累计3次及以上严厉警告员工明细",
                "branch_risk": "网点员工风险预警"
            },
            'english': {
                "details": "details",
                "2x_stern": "Staff_with_2x_Stern_Reminders",       
                "3x_stern": "Staff_with_3x+_Stern_Reminders",  
                "branch_risk": "Branch_Staff_Risk_Alert"
            }
        }
        
        # --- Field mappings for English sheets ---
        self.FIELD_MAPPINGS = {
            'chinese': {
                "满2次严厉警告员工人数": "满2次严厉警告员工人数",
                "超3次及以上严厉警告员工人数": "超3次及以上严厉警告员工人数"
            },
            'english': {
                "满2次严厉警告员工人数": "EmployeeCount_with_2x_Stern_Reminders",
                "超3次及以上严厉警告员工人数": "EmployeeCount_with_3x+_Stern_Reminders"
            }
        }

        # 1. 为新增的复选框创建控制变量
        self.use_filename_as_subject_var = tk.BooleanVar(value=False)

        self.create_widgets()

    def create_widgets(self):
        # --- 创建带滚动条功能的主画布 ---
        self.canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.bind_mousewheel()

        # --- 新的根布局框架 (在可滚动区域内) ---
        root_frame = ttk.Frame(self.scrollable_frame, padding="10")
        root_frame.pack(fill=tk.BOTH, expand=True)

        # --- 右侧栏: 快捷操作 (如配置文件上传) ---
        right_column_frame = ttk.Frame(root_frame)
        right_column_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        
        config_frame = ttk.LabelFrame(right_column_frame, text="快捷操作", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(config_frame, text="从文件加载配置", command=self.load_configuration_file).pack(pady=10, fill=tk.X, ipady=4)

        # --- 左侧栏: 主流程配置 ---
        left_column_frame = ttk.Frame(root_frame)
        left_column_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # --- 1. 发件人设置 ---
        sender_frame = ttk.LabelFrame(left_column_frame, text="1. 发件人设置 (推荐: 腾讯企业邮箱)", padding="10")
        sender_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(sender_frame, text="SMTP服务器:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.smtp_server_var = tk.StringVar(value="smtp.exmail.qq.com")
        ttk.Entry(sender_frame, textvariable=self.smtp_server_var, width=30).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(sender_frame, text="SMTP端口:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.smtp_port_var = tk.StringVar(value="465")
        ttk.Entry(sender_frame, textvariable=self.smtp_port_var, width=10).grid(row=1, column=3, padx=5, pady=5)

        ttk.Label(sender_frame, text="发件人邮箱:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.sender_email_var = tk.StringVar()
        ttk.Entry(sender_frame, textvariable=self.sender_email_var, width=30).grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(sender_frame, text="邮箱授权码:").grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.password_var = tk.StringVar()
        ttk.Entry(sender_frame, textvariable=self.password_var, show="*", width=20).grid(row=2, column=3, padx=5, pady=5)
        
        # --- 2. 数据与文件选择 ---
        data_frame = ttk.LabelFrame(left_column_frame, text="2. 数据与文件选择", padding="10")
        data_frame.pack(fill=tk.X, pady=5)

        ttk.Label(data_frame, text="选择原始数据文件 (Excel):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.source_file_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.source_file_var, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(data_frame, text="浏览...", command=self.select_source_file).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(data_frame, text="选择用于拆分的字段:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.split_column_var = tk.StringVar()
        self.split_column_combo = ttk.Combobox(data_frame, textvariable=self.split_column_var, width=57, state="readonly")
        self.split_column_combo.grid(row=1, column=1, padx=5, pady=5)
        self.split_column_combo.bind('<<ComboboxSelected>>', self.on_split_column_selected)

        ttk.Label(data_frame, text="全英文处理配置 - 选择需要英文处理的拆分值:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        
        self.english_values_frame = ttk.Frame(data_frame)
        self.english_values_frame.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        
        self.english_checkboxes = {}
        self.english_checkbox_widgets = {}

        ttk.Label(data_frame, text="选择邮箱映射关系文件 (Excel):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.mapping_file_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.mapping_file_var, width=60).grid(row=4, column=1, padx=5, pady=5)
        ttk.Button(data_frame, text="浏览...", command=self.select_mapping_file).grid(row=4, column=2, padx=5, pady=5)

        ttk.Label(data_frame, text="选择拆分后表格保存位置:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.save_dir_var = tk.StringVar()
        ttk.Entry(data_frame, textvariable=self.save_dir_var, width=60).grid(row=5, column=1, padx=5, pady=5)
        ttk.Button(data_frame, text="浏览...", command=self.select_save_directory).grid(row=5, column=2, padx=5, pady=5)

        # --- 3. 邮件内容配置 ---
        content_frame = ttk.LabelFrame(left_column_frame, text="3. 邮件内容配置", padding="10")
        content_frame.pack(fill=tk.X, pady=5)

        # 2. 修改邮件主题行，增加复选框
        ttk.Label(content_frame, text="邮件主题 (前缀):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.subject_var = tk.StringVar()
        self.subject_entry = ttk.Entry(content_frame, textvariable=self.subject_var, width=60)
        self.subject_entry.grid(row=0, column=1, padx=5, pady=5)
        
        self.subject_checkbox = ttk.Checkbutton(
            content_frame,
            text="使用源文件名作为前缀",
            variable=self.use_filename_as_subject_var,
            command=self.toggle_subject_entry_state
        )
        self.subject_checkbox.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        ttk.Label(content_frame, text="抄送人员 (多个用英文分号';'隔开):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.cc_var = tk.StringVar()
        ttk.Entry(content_frame, textvariable=self.cc_var, width=80).grid(row=1, column=1, columnspan=3, padx=5, pady=5)

        ttk.Label(content_frame, text="中文邮件正文前缀:").grid(row=2, column=0, padx=5, pady=5, sticky="nw")
        self.chinese_prefix_text = tk.Text(content_frame, height=4, width=80)
        self.chinese_prefix_text.grid(row=2, column=1, columnspan=3, padx=5, pady=5)

        ttk.Label(content_frame, text="中文邮件正文后缀:").grid(row=3, column=0, padx=5, pady=5, sticky="nw")
        self.chinese_suffix_text = tk.Text(content_frame, height=4, width=80)
        self.chinese_suffix_text.grid(row=3, column=1, columnspan=3, padx=5, pady=5)

        ttk.Label(content_frame, text="英文邮件正文前缀:").grid(row=4, column=0, padx=5, pady=5, sticky="nw")
        self.english_prefix_text = tk.Text(content_frame, height=4, width=80)
        self.english_prefix_text.grid(row=4, column=1, columnspan=3, padx=5, pady=5)

        ttk.Label(content_frame, text="英文邮件正文后缀:").grid(row=5, column=0, padx=5, pady=5, sticky="nw")
        self.english_suffix_text = tk.Text(content_frame, height=4, width=80)
        self.english_suffix_text.grid(row=5, column=1, columnspan=3, padx=5, pady=5)

        # --- 4. 执行与进度 ---
        exec_frame = ttk.LabelFrame(left_column_frame, text="4. 执行与进度", padding="10")
        exec_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.start_button = ttk.Button(exec_frame, text="一键执行 (包含中英文智能处理)", command=self.start_sending_thread)
        self.start_button.pack(pady=10, ipady=4)

        self.progress = ttk.Progressbar(exec_frame, orient="horizontal", length=750, mode="determinate")
        self.progress.pack(pady=5)

        self.status_log = tk.Text(exec_frame, height=12, state="disabled")
        self.status_log.pack(fill=tk.BOTH, expand=True, pady=5)

    # 3. 新增方法，用于切换主题输入框的状态
    def toggle_subject_entry_state(self):
        """当复选框被点击时，启用或禁用主题前缀输入框"""
        if self.use_filename_as_subject_var.get():
            self.subject_entry.config(state="disabled")
            self.subject_var.set("") # 清空内容
        else:
            self.subject_entry.config(state="normal")

    def bind_mousewheel(self):
        """Bind mousewheel events to canvas for scrolling"""
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            self.canvas.unbind_all("<MouseWheel>")
        
        self.canvas.bind('<Enter>', _bind_to_mousewheel)
        self.canvas.bind('<Leave>', _unbind_from_mousewheel)

    def log(self, message):
        self.status_log.config(state="normal")
        self.status_log.insert(tk.END, message + "\n")
        self.status_log.see(tk.END)
        self.status_log.config(state="disabled")
        self.update_idletasks()

    def load_configuration_file(self):
        filepath = filedialog.askopenfilename(
            title="选择配置文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if not filepath:
            return

        try:
            if filepath.endswith('.csv'):
                df = pd.read_csv(filepath, header=None)
            else:
                df = pd.read_excel(filepath, header=None)
            
            config_dict = pd.Series(df.iloc[:, 1].values, index=df.iloc[:, 0]).dropna().to_dict()

            UI_VARS_MAP = {
                "smtp_server": self.smtp_server_var,
                "smtp_port": self.smtp_port_var,
                "sender_email": self.sender_email_var,
                "password": self.password_var,
                "subject_prefix": self.subject_var,
                "cc_recipients": self.cc_var,
                "chinese_prefix": self.chinese_prefix_text,
                "chinese_suffix": self.chinese_suffix_text,
                "english_prefix": self.english_prefix_text,
                "english_suffix": self.english_suffix_text
            }

            for key, value in config_dict.items():
                widget = UI_VARS_MAP.get(key)
                if widget:
                    if isinstance(widget, tk.StringVar):
                        widget.set(str(value))
                    elif isinstance(widget, tk.Text):
                        widget.delete("1.0", tk.END)
                        widget.insert("1.0", str(value))
            
            self.log("配置文件加载成功。")
            messagebox.showinfo("成功", "已从文件成功加载配置。")

        except Exception as e:
            self.log(f"错误: 加载配置文件失败: {e}")
            messagebox.showerror("错误", f"无法读取或解析配置文件。\n请确保文件是两列（Key, Value）并且格式正确。\n\n错误详情: {e}")

    def _get_column_mappings(self, df_columns):
        """Intelligently map dataframe columns to standard internal names."""
        self.log("开始智能列名映射...")
        self.COLUMN_MAP = {}
        df_columns_lower = {col.lower().strip(): col for col in df_columns}

        for key, possible_names in self.KEY_COLS.items():
            for name in possible_names:
                found = False
                for lower_col, original_col in df_columns_lower.items():
                    if name in lower_col:
                        self.COLUMN_MAP[key] = original_col
                        self.log(f"  ✓ 映射成功: '{key}' -> '{original_col}'")
                        found = True
                        break
                if found:
                    break
        
        for crit_col in self.CRITICAL_COLS:
            if crit_col not in self.COLUMN_MAP:
                raise ValueError(f"关键列缺失！程序无法在文件中找到代表'{crit_col}'的列。请确保文件包含一个标题含有以下关键词之一的列: {self.KEY_COLS[crit_col]}")

        self.log("智能列名映射完成。")
        return self.COLUMN_MAP

    def select_source_file(self):
        filepath = filedialog.askopenfilename(title="选择原始数据文件", filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if filepath:
            self.source_file_var.set(filepath)
            try:
                if filepath.endswith('.csv'):
                    df = pd.read_csv(filepath)
                else:
                    df = pd.read_excel(filepath)
                
                self.split_column_combo['values'] = list(df.columns)
                self.log(f"成功加载原始文件: {os.path.basename(filepath)}")
                self.log("请在下拉菜单中选择用于拆分的字段。")
                
                self.source_df = df
                self._get_column_mappings(df.columns)
            except Exception as e:
                messagebox.showerror("错误", f"无法读取文件表头或映射列: {e}")
                self.log(f"错误: 无法读取或映射文件 {os.path.basename(filepath)}. 错误详情: {e}")

    def on_split_column_selected(self, event=None):
        """当选择拆分字段时，更新英文处理选项"""
        if not hasattr(self, 'source_df') or self.source_df is None:
            return
        
        split_column = self.split_column_var.get()
        if not split_column:
            return
        
        try:
            unique_values = sorted(self.source_df[split_column].dropna().unique())
            self.split_field_values = [str(val) for val in unique_values]
            
            for widget in self.english_checkbox_widgets.values():
                widget.destroy()
            self.english_checkboxes.clear()
            self.english_checkbox_widgets.clear()
            
            self.log(f"已检测到拆分字段 '{split_column}' 的 {len(self.split_field_values)} 个唯一值")
            
            max_cols = 4
            for i, value in enumerate(self.split_field_values):
                row = i // max_cols
                col = i % max_cols
                
                var = tk.BooleanVar()
                self.english_checkboxes[value] = var
                
                checkbox = ttk.Checkbutton(
                    self.english_values_frame, 
                    text=f"{value}", 
                    variable=var,
                    command=self.update_english_processing_values
                )
                checkbox.grid(row=row, column=col, padx=10, pady=2, sticky="w")
                self.english_checkbox_widgets[value] = checkbox
            
            self.log("可选择需要全英文处理的拆分值（可选择多个或不选）")
            
        except Exception as e:
            self.log(f"处理拆分字段时出错: {e}")

    def update_english_processing_values(self):
        """更新需要英文处理的值列表"""
        self.english_processing_values = set()
        for value, var in self.english_checkboxes.items():
            if var.get():
                self.english_processing_values.add(value)
        
        if self.english_processing_values:
            selected_values = ', '.join(sorted(self.english_processing_values))
            self.log(f"已选择进行全英文处理的值: {selected_values}")
        else:
            self.log("当前未选择任何值进行全英文处理")

    def select_mapping_file(self):
        filepath = filedialog.askopenfilename(title="选择邮箱映射关系文件", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.mapping_file_var.set(filepath)
            self.log(f"已选择邮箱映射文件: {os.path.basename(filepath)}")

    def select_save_directory(self):
        dirpath = filedialog.askdirectory(title="选择保存位置")
        if dirpath:
            self.save_dir_var.set(dirpath)
            self.log(f"拆分文件将保存至: {dirpath}")
            
    def start_sending_thread(self):
        self.start_button.config(state="disabled")
        self.status_log.config(state="normal")
        self.status_log.delete(1.0, tk.END)
        self.status_log.config(state="disabled")
        
        self.update_english_processing_values()
        if self.english_processing_values:
            self.log(f"最终确认需要全英文处理的字段值: {', '.join(sorted(self.english_processing_values))}")
        else:
            self.log("未选择任何值进行全英文处理，将使用默认中文模式")
        
        processing_thread = threading.Thread(target=self.process_and_send_emails)
        processing_thread.daemon = True 
        processing_thread.start()

    def preprocess_data(self, df):
        """
        数据预处理和标准化，根据用户要求调整警告类型处理逻辑
        """
        self.log("开始数据预处理和标准化...")
        df = df.copy() 

        warning_type_col = self.COLUMN_MAP.get('warning_type')
        if warning_type_col:
            self.log(f"正在标准化 '{warning_type_col}' 列...")
            df[warning_type_col] = df[warning_type_col].astype(str)

            # --- 新的警告类型标准化逻辑，保留原始语言 ---
            def standardize_warning_type(value):
                value = str(value).strip()
                # 如果值包含中文字符，则标准化为中文
                if re.search('[\u4e00-\u9fff]', value):
                    if "严厉" in value:
                        return "严厉警告"
                    elif "口述" in value:
                        return "口述警告"
                    # 如果是中文，但不是警告类型，则原样返回
                    return value 
                # 如果值不包含中文字符，则标准化为英文
                else:
                    if "stern" in value.lower():
                        return "Stern Reminder"
                    elif "verbal" in value.lower():
                        return "Verbal Warning"
                    # 如果是英文，但不是警告类型，则原样返回
                    return value
            
            df[warning_type_col] = df[warning_type_col].apply(standardize_warning_type)
            self.log("  - '警告类型'字段值已根据原语言类型（中/英）完成统一。")

        sending_status_col = self.COLUMN_MAP.get('sending_status')
        if sending_status_col:
            self.log(f"正在标准化 '{sending_status_col}' 列...")
            df[sending_status_col] = df[sending_status_col].astype(str)

            df.loc[df[sending_status_col].str.contains("待发送", na=False), sending_status_col] = "已发送"
            df.loc[df[sending_status_col].str.contains("Pending", na=False, case=False), sending_status_col] = "Has been sent"
            self.log("  - '发送状态'字段值已根据关键词完成统一。")

        self.log("数据预处理和标准化完成。")
        return df

    def is_english_processing_required(self, area_name):
        return str(area_name).strip() in self.english_processing_values

    def generate_sheet_names(self, is_english=False):
        return self.SHEET_NAMES['english' if is_english else 'chinese']

    def count_warnings_per_employee(self, df):
        """
        核心逻辑：使用映射关系进行统一的警告统计
        """
        id_col = self.COLUMN_MAP.get('id')
        warning_type_col = self.COLUMN_MAP.get('warning_type')
        
        if not id_col or not warning_type_col:
            self.log("错误: 无法统计警告，因为员工ID或警告类型列未被映射。")
            return {}

        # 映射警告类型为统一的英文表示以便进行统计
        mapping_for_stats = {
            "严厉警告": "Stern Reminder", 
            "口述警告": "Verbal Warning"
        }
        temp_warning_types = df[warning_type_col].replace(mapping_for_stats)
        
        # 将原始中文和英文数据统一计数
        warning_counts_df = df.groupby([id_col, temp_warning_types]).size().unstack(fill_value=0)
        
        result = {}
        # 确保每个员工的两种警告类型都有记录，即使为0
        unique_ids = df[id_col].unique()
        for emp_id in unique_ids:
            counts = warning_counts_df.loc[emp_id] if emp_id in warning_counts_df.index else {}
            result[emp_id] = {
                "Stern Reminder": counts.get("Stern Reminder", 0),
                "Verbal Warning": counts.get("Verbal Warning", 0)
            }
        return result

    def generate_warning_analysis_sheets(self, df_split, all_employees_warning_counts, is_english=False):
        sheets_data = {}
        sheet_names = self.generate_sheet_names(is_english)
        field_mappings = self.FIELD_MAPPINGS['english' if is_english else 'chinese']
        
        id_col = self.COLUMN_MAP.get('id')
        branch_col = self.COLUMN_MAP.get('branch')
        
        base_columns_keys = ['area', 'district', 'branch', 'ops', 'position', 'id', 'name', 'status', 'employment_type']
        available_base_columns = [self.COLUMN_MAP[key] for key in base_columns_keys if key in self.COLUMN_MAP]
        
        unique_employees_in_split = df_split.drop_duplicates(subset=[id_col])

        two_stern_employees = []
        three_plus_stern_employees = []
        for _, emp_data in unique_employees_in_split.iterrows():
            emp_id = emp_data[id_col]
            counts = all_employees_warning_counts.get(emp_id, {"Stern Reminder": 0, "Verbal Warning": 0})
            
            row_data = emp_data[available_base_columns].to_dict()
            # 增加统计列，并根据模式命名
            if is_english:
                row_data["Stern Reminder"] = counts["Stern Reminder"]
                row_data["Verbal Warning"] = counts["Verbal Warning"]
            else:
                row_data["严厉警告 Stern Reminder"] = counts["Stern Reminder"]
                row_data["口述警告 Verbal Warning"] = counts["Verbal Warning"]
            
            if counts["Stern Reminder"] == 2:
                two_stern_employees.append(row_data)
            elif counts["Stern Reminder"] >= 3:
                three_plus_stern_employees.append(row_data)
        
        if two_stern_employees:
            df_2x = pd.DataFrame(two_stern_employees)
            df_2x = df_2x.sort_values(df_2x.columns[-2], ascending=False)
            sheets_data[sheet_names["2x_stern"]] = df_2x
            self.log(f"分析生成 '{sheet_names['2x_stern']}': {len(two_stern_employees)} 名员工")
        
        if three_plus_stern_employees:
            df_3x = pd.DataFrame(three_plus_stern_employees)
            df_3x = df_3x.sort_values(df_3x.columns[-2], ascending=False)
            sheets_data[sheet_names["3x_stern"]] = df_3x
            self.log(f"分析生成 '{sheet_names['3x_stern']}': {len(three_plus_stern_employees)} 名员工")
        
        if branch_col:
            branch_risk_data = []
            branches = df_split[branch_col].dropna().unique()
            for branch in branches:
                branch_employees_df = df_split[df_split[branch_col] == branch]
                branch_employee_ids = branch_employees_df[id_col].unique()
                count_2x = sum(1 for emp_id in branch_employee_ids if all_employees_warning_counts.get(emp_id, {}).get("Stern Reminder", 0) == 2)
                count_3x_plus = sum(1 for emp_id in branch_employee_ids if all_employees_warning_counts.get(emp_id, {}).get("Stern Reminder", 0) >= 3)
                if count_2x > 0 or count_3x_plus > 0:
                    branch_info = branch_employees_df.iloc[0]
                    risk_row = {
                        self.COLUMN_MAP.get('area', 'Area'): branch_info.get(self.COLUMN_MAP.get('area')),
                        self.COLUMN_MAP.get('district', 'District'): branch_info.get(self.COLUMN_MAP.get('district')),
                        branch_col: branch,
                        self.COLUMN_MAP.get('ops', 'OPS'): branch_info.get(self.COLUMN_MAP.get('ops')),
                        field_mappings["满2次严厉警告员工人数"]: count_2x,
                        field_mappings["超3次及以上严厉警告员工人数"]: count_3x_plus
                    }
                    branch_risk_data.append(risk_row)
            
            if branch_risk_data:
                branch_risk_df = pd.DataFrame(branch_risk_data)
                branch_risk_df["_总风险人数"] = (
                    branch_risk_df[field_mappings["满2次严厉警告员工人数"]].fillna(0) + 
                    branch_risk_df[field_mappings["超3次及以上严厉警告员工人数"]].fillna(0)
                )
                branch_risk_df = branch_risk_df.sort_values("_总风险人数", ascending=False).drop("_总风险人数", axis=1)
                sheets_data[sheet_names["branch_risk"]] = branch_risk_df
                self.log(f"分析生成 '{sheet_names['branch_risk']}': {len(branch_risk_df)} 个网点")
        
        return sheets_data

    def generate_statistics_summary(self, df_split, all_employees_warning_counts, is_english=False):
        id_col = self.COLUMN_MAP.get('id')
        branch_col = self.COLUMN_MAP.get('branch')
        status_col = self.COLUMN_MAP.get('status')
        
        unique_employees = df_split.drop_duplicates(subset=[id_col])
        
        count_2x_total = 0
        count_3x_plus_total = 0
        status_2x_count = {}
        status_3x_plus_count = {}
        
        for _, emp_data in unique_employees.iterrows():
            emp_id = emp_data[id_col]
            counts = all_employees_warning_counts.get(emp_id, {"Stern Reminder": 0, "Verbal Warning": 0})
            status = emp_data.get(status_col, "未知" if not is_english else "Unknown")
            
            if counts["Stern Reminder"] == 2:
                count_2x_total += 1
                status_2x_count[status] = status_2x_count.get(status, 0) + 1
            elif counts["Stern Reminder"] >= 3:
                count_3x_plus_total += 1
                status_3x_plus_count[status] = status_3x_plus_count.get(status, 0) + 1
        
        branch_risk_details = {}
        if branch_col:
            unique_branches = df_split[branch_col].dropna().unique()
            for branch in unique_branches:
                details = {
                    'total': 0, '2x_count': 0, '3x_plus_count': 0,
                    '2x_status_breakdown': {}, '3x_plus_status_breakdown': {}
                }
                branch_employees = df_split[df_split[branch_col] == branch]
                branch_unique_employees = branch_employees.drop_duplicates(subset=[id_col])

                for _, emp_data in branch_unique_employees.iterrows():
                    emp_id = emp_data[id_col]
                    counts = all_employees_warning_counts.get(emp_id, {})
                    status = emp_data.get(status_col, "未知" if not is_english else "Unknown")
                    
                    if counts.get("Stern Reminder", 0) == 2:
                        details['2x_count'] += 1
                        breakdown = details['2x_status_breakdown']
                        breakdown[status] = breakdown.get(status, 0) + 1
                    elif counts.get("Stern Reminder", 0) >= 3:
                        details['3x_plus_count'] += 1
                        breakdown = details['3x_plus_status_breakdown']
                        breakdown[status] = breakdown.get(status, 0) + 1
                
                details['total'] = details['2x_count'] + details['3x_plus_count']
                if details['total'] > 0:
                    branch_risk_details[branch] = details
        
        top_5_branches = sorted(branch_risk_details.items(), key=lambda x: x[1]['total'], reverse=True)[:5]
        
        if is_english:
            summary = f"""=== WARNING STATISTICS SUMMARY ===

Total Employees Analysis:
- Employees with exactly 2 Stern Reminders: {count_2x_total}
- Employees with 3+ Stern Reminders: {count_3x_plus_total}

Work Status Distribution (2 Stern Reminders):
{self._format_status_breakdown(status_2x_count, is_english)}

Work Status Distribution (3+ Stern Reminders):
{self._format_status_breakdown(status_3x_plus_count, is_english)}

Branches Involved: {len(branch_risk_details)}

Top 5 High-Risk Branches:
{self._format_top_branches_detailed(top_5_branches, is_english)}
==============================="""
        else:
            summary = f"""=== 警告统计总结 ===

员工总体分析:
- 累计2次严厉警告员工: {count_2x_total}人
- 累计3次及以上严厉警告员工: {count_3x_plus_total}人

在职状态分布 (2次严厉警告):
{self._format_status_breakdown(status_2x_count, is_english)}

在职状态分布 (3次及以上严厉警告):
{self._format_status_breakdown(status_3x_plus_count, is_english)}

涉及网点数量: {len(branch_risk_details)}个

高风险网点排名前5:
{self._format_top_branches_detailed(top_5_branches, is_english)}
===================="""
        return summary.strip()

    def _format_status_breakdown(self, status_count, is_english=False):
        if not status_count:
            return "- 无数据" if not is_english else "- No data"
        
        result = []
        for status, count in status_count.items():
            result.append(f"- {status}: {count}{'人' if not is_english else ''}")
        return "\n".join(result)

    def _format_status_breakdown_inline(self, breakdown_dict, is_english=False):
        if not breakdown_dict:
            return ""
        
        separator = ", "
        if is_english:
            items = [f"{status}: {count}" for status, count in breakdown_dict.items()]
        else:
            items = [f"{status}: {count}人" for status, count in breakdown_dict.items()]
        
        return separator.join(items)

    def _format_top_branches_detailed(self, top_branches, is_english=False):
        if not top_branches:
            return "- 无数据" if not is_english else "- No data"
        
        result = []
        for i, (branch, details) in enumerate(top_branches, 1):
            if is_english:
                branch_str = f"{i}. {branch} - Total Risk: {details['total']}"
                if details['3x_plus_count'] > 0:
                    status_dist_3x = self._format_status_breakdown_inline(details['3x_plus_status_breakdown'], is_english)
                    branch_str += f"\n   - Employees with 3+ Stern Reminders: {details['3x_plus_count']}; Status distribution: {status_dist_3x}"
                if details['2x_count'] > 0:
                    status_dist_2x = self._format_status_breakdown_inline(details['2x_status_breakdown'], is_english)
                    branch_str += f"\n   - Employees with 2 Stern Reminders: {details['2x_count']}; Status distribution: {status_dist_2x}"
            else:
                branch_str = f"{i}. {branch} - 总风险人数: {details['total']}人"
                if details['3x_plus_count'] > 0:
                    status_dist_3x = self._format_status_breakdown_inline(details['3x_plus_status_breakdown'], is_english)
                    branch_str += f"\n   - 累计3次及以上严厉警告员工: {details['3x_plus_count']}人；状态分布: {status_dist_3x}"
                if details['2x_count'] > 0:
                    status_dist_2x = self._format_status_breakdown_inline(details['2x_status_breakdown'], is_english)
                    branch_str += f"\n   - 累计2次严厉警告员工: {details['2x_count']}人；状态分布: {status_dist_2x}"
            result.append(branch_str)
        return "\n".join(result)
    
    def create_multi_sheet_excel(self, df_split, file_path, area_name, all_employees_warning_counts):
        is_english = self.is_english_processing_required(area_name)
        sheet_names = self.generate_sheet_names(is_english)
        
        analysis_sheets = self.generate_warning_analysis_sheets(df_split, all_employees_warning_counts, is_english)
        
        ordered_sheets = []
        
        if sheet_names["branch_risk"] in analysis_sheets:
            ordered_sheets.append((sheet_names["branch_risk"], analysis_sheets[sheet_names["branch_risk"]]))
        
        # 修正逻辑：使用翻译后的键名访问字典
        for sheet_key in ["3x_stern", "2x_stern"]:
            translated_sheet_name = sheet_names[sheet_key]
            if translated_sheet_name in analysis_sheets:
                ordered_sheets.append((translated_sheet_name, analysis_sheets[translated_sheet_name]))
        
        ordered_sheets.append((sheet_names["details"], df_split))
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, sheet_df in ordered_sheets:
                sheet_df_copy = sheet_df.copy()

                stat_cols_to_clear_zero = [
                    "严厉警告 Stern Reminder", 
                    "口述警告 Verbal Warning",
                    "Stern Reminder",
                    "Verbal Warning",
                    self.FIELD_MAPPINGS['chinese']["满2次严厉警告员工人数"],
                    self.FIELD_MAPPINGS['english']["满2次严厉警告员工人数"],
                    self.FIELD_MAPPINGS['chinese']["超3次及以上严厉警告员工人数"],
                    self.FIELD_MAPPINGS['english']["超3次及以上严厉警告员工人数"]
                ]
                
                # 在分析表中将值为0的统计列置为空，以便更清晰地展示
                if sheet_name != sheet_names["details"]:
                    for col in stat_cols_to_clear_zero:
                        if col in sheet_df_copy.columns:
                            sheet_df_copy[col] = sheet_df_copy[col].replace(0, None)
                
                sheet_df_copy.to_excel(writer, sheet_name=sheet_name, index=False)
        
        self.log(f"多sheet Excel文件已生成: {os.path.basename(file_path)} ({'英文模式' if is_english else '中文模式'})")

    def generate_email_content(self, area_name, df_split, all_employees_warning_counts):
        is_english = self.is_english_processing_required(area_name)
        
        if is_english:
            prefix = self.english_prefix_text.get("1.0", tk.END).strip()
            suffix = self.english_suffix_text.get("1.0", tk.END).strip()
        else:
            prefix = self.chinese_prefix_text.get("1.0", tk.END).strip()
            suffix = self.chinese_suffix_text.get("1.0", tk.END).strip()
        
        summary = self.generate_statistics_summary(df_split, all_employees_warning_counts, is_english)
        
        email_body = f"{prefix}\n\n{summary}\n\n{suffix}"
        
        return email_body.strip()

    def process_and_send_emails(self):
        try:
            self.log("开始执行任务，正在检查参数...")

            # 首先根据复选框状态确定邮件主题前缀
            subject_prefix = ""
            if self.use_filename_as_subject_var.get():
                source_file_path = self.source_file_var.get()
                if not source_file_path:
                    raise ValueError("请先选择源数据文件，再勾选“使用源文件名作为前缀”。")
                # 提取不含扩展名的文件名
                subject_prefix = os.path.splitext(os.path.basename(source_file_path))[0]
                self.log(f"已设置邮件主题前缀为源文件名: '{subject_prefix}'")
            else:
                subject_prefix = self.subject_var.get()

            params = {
                "source_file": self.source_file_var.get(),
                "split_column": self.split_column_var.get(),
                "mapping_file": self.mapping_file_var.get(),
                "save_dir": self.save_dir_var.get(),
                "smtp_server": self.smtp_server_var.get(),
                "smtp_port": self.smtp_port_var.get(),
                "sender_email": self.sender_email_var.get(),
                "password": self.password_var.get(),
                "subject_prefix": subject_prefix, # 使用最终确定的前缀
                "cc_recipients": [cc.strip() for cc in self.cc_var.get().split(';') if cc.strip()]
            }
            
            # 在参数确定后进行校验
            required_fields = ["source_file", "split_column", "mapping_file", "save_dir", 
                             "smtp_server", "smtp_port", "sender_email", "password", "subject_prefix"]
            for field in required_fields:
                if not params[field]:
                    if field == "subject_prefix":
                        raise ValueError("邮件主题前缀不能为空！请填写或选择使用源文件名。")
                    raise ValueError(f"必填项 '{field}' 不能为空！请检查配置。")
            
            chinese_prefix = self.chinese_prefix_text.get("1.0", tk.END).strip()
            chinese_suffix = self.chinese_suffix_text.get("1.0", tk.END).strip()
            english_prefix = self.english_prefix_text.get("1.0", tk.END).strip()
            english_suffix = self.english_suffix_text.get("1.0", tk.END).strip()
            
            # 至少一种语言的邮件模板需要填写
            if not (chinese_prefix or chinese_suffix) and not (english_prefix or english_suffix):
                raise ValueError("中文或英文邮件模板（前缀或后缀）至少需要填写一个！")
            
            self.log("参数校验通过。")

            self.log("正在读取原始数据文件...")
            source_file = params["source_file"]
            if source_file.endswith('.csv'):
                df_source = pd.read_csv(source_file)
            else:
                df_source = pd.read_excel(source_file)
            self.log(f"原始数据文件包含 {len(df_source)} 行数据。")
            
            self._get_column_mappings(df_source.columns)
            
            df_source_preprocessed = self.preprocess_data(df_source)
            
            self.log("正在对所有员工进行警告次数全局统计...")
            all_employees_warning_counts = self.count_warnings_per_employee(df_source_preprocessed)
            self.log(f"全局统计完成，共分析了 {len(all_employees_warning_counts)} 名不重复员工。")
            
            self.log("正在读取邮箱映射文件...")
            df_mapping = pd.read_excel(params["mapping_file"])
            mapping_dict = pd.Series(df_mapping.iloc[:, 1].values, index=df_mapping.iloc[:, 0]).to_dict()
            self.log("邮箱映射关系加载成功。")

            split_values = df_source_preprocessed[params["split_column"]].unique()
            total_tasks = len(split_values)
            self.progress['maximum'] = total_tasks
            self.log(f"检测到需要处理 {total_tasks} 个不同的拆分值。")

            for i, value in enumerate(split_values):
                self.log("-" * 60)
                self.log(f"正在处理: [{value}] ({i+1}/{total_tasks})")

                df_split = df_source_preprocessed[df_source_preprocessed[params["split_column"]] == value].copy()
                self.log(f"拆分后的数据包含 {len(df_split)} 行。")
                
                is_english_processing = self.is_english_processing_required(value)
                processing_mode = "英文模式" if is_english_processing else "中文模式"
                self.log(f"当前处理模式: {processing_mode}")
                
                attachment_filename = f"{params['subject_prefix']}_{value}.xlsx"
                attachment_path = os.path.join(params["save_dir"], attachment_filename)
                
                self.create_multi_sheet_excel(df_split, attachment_path, value, all_employees_warning_counts)

                recipient_email = mapping_dict.get(value)
                if not recipient_email:
                    self.log(f"警告: 在映射表中未找到值 '{value}' 对应的邮箱，跳过此项。")
                    self.progress['value'] = i + 1
                    continue
                self.log(f"查找到收件人: {recipient_email}")

                email_body = self.generate_email_content(value, df_split, all_employees_warning_counts)
                
                final_subject = f"{params['subject_prefix']}_{value}"
                msg = MIMEMultipart()
                msg['From'] = params["sender_email"]
                msg['To'] = recipient_email
                msg['Subject'] = final_subject
                if params["cc_recipients"]:
                    msg['Cc'] = ";".join(params["cc_recipients"])
                msg.attach(MIMEText(email_body, 'plain', 'utf-8'))

                try:
                    with open(attachment_path, "rb") as f:
                        part = MIMEApplication(f.read(), Name=attachment_filename)
                    part['Content-Disposition'] = f'attachment; filename="{attachment_filename}"'
                    msg.attach(part)
                    self.log(f"多sheet Excel附件已添加: {attachment_filename}")
                except Exception as attach_error:
                    self.log(f"添加附件时出错: {attach_error}")
                    continue

                all_recipients = [recipient_email] + params["cc_recipients"]
                self.log(f"正在连接SMTP服务器: {params['smtp_server']}:{params['smtp_port']}...")
                try:
                    with smtplib.SMTP_SSL(params["smtp_server"], int(params["smtp_port"])) as server:
                        server.login(params["sender_email"], params["password"])
                        self.log("服务器登录成功，准备发送邮件...")
                        server.sendmail(params["sender_email"], all_recipients, msg.as_string())
                        cc_info = f"抄送: {';'.join(params['cc_recipients'])}" if params["cc_recipients"] else "无抄送"
                        self.log(f"邮件已成功发送至: {recipient_email} ({cc_info}) - 处理模式: {processing_mode}")
                except Exception as email_error:
                    self.log(f"邮件发送失败: {email_error}")
                    continue

                self.progress['value'] = i + 1

            self.log("-" * 60)
            self.log("所有任务执行完毕！")
            messagebox.showinfo("完成", "所有邮件已成功发送！包含中英文智能处理和统计分析功能。")

        except Exception as e:
            self.log(f"发生致命错误: {e}")
            messagebox.showerror("执行出错", f"任务中断，请检查日志窗口中的错误信息。\n\n错误详情: {e}")
        finally:
            self.start_button.config(state="normal")

if __name__ == "__main__":
    app = EmailSenderApp()
    app.mainloop()