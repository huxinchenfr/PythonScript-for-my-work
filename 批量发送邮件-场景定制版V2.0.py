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
    An automated tool for batch sending emails, specifically for processing historical warning letters.
    """
    def __init__(self):
        super().__init__()

        # --- Language Configuration ---
        self.LANG = {
            'zh': {
                "title": "自动化邮件批量发送工具-历史警告信数据处理专版",
                "quick_actions": "快捷操作",
                "load_config": "从文件加载配置",
                "lang_switch": "语言 (Language)",
                "sender_settings": "1. 发件人设置 (推荐: 腾讯企业邮箱)",
                "smtp_server": "SMTP服务器:",
                "smtp_port": "SMTP端口:",
                "sender_email": "发件人邮箱:",
                "email_auth_code": "邮箱授权码:",
                "data_files": "2. 数据与文件选择",
                "select_source_file": "选择原始数据文件 (Excel):",
                "browse": "浏览...",
                "select_split_field": "选择用于拆分的字段:",
                "english_config": "全英文处理配置 - 选择需要英文处理的拆分值:",
                "select_mapping_file": "选择邮箱映射关系文件 (Excel):",
                "select_save_location": "选择拆分后表格保存位置:",
                "email_content": "3. 邮件内容配置",
                "subject_prefix": "邮件主题 (前缀):",
                "use_filename_as_prefix": "使用源文件名作为前缀",
                "cc_recipients": "抄送人员 (多个用英文分号';'隔开):",
                "chinese_prefix": "中文邮件正文前缀:",
                "chinese_suffix": "中文邮件正文后缀:",
                "english_prefix": "英文邮件正文前缀:",
                "english_suffix": "英文邮件正文后缀:",
                "execution_progress": "4. 执行与进度",
                "execute_button": "一键执行 (包含中英文智能处理)",
                "success_title": "成功",
                "config_loaded_msg": "已从文件成功加载配置。",
                "error_title": "错误",
                "config_load_error_msg": "无法读取或解析配置文件。\n请确保文件是两列（Key, Value）并且格式正确。\n\n错误详情: {}",
                "file_read_error_msg": "无法读取文件表头或映射列: {}",
                "all_tasks_complete_title": "完成",
                "all_tasks_complete_msg": "所有邮件已成功发送！包含中英文智能处理和统计分析功能。",
                "execution_error_title": "执行出错",
                "execution_error_msg": "任务中断，请检查日志窗口中的错误信息。\n\n错误详情: {}"
            },
            'en': {
                "title": "Automated Bulk Email Sender - Warning Letter Edition",
                "quick_actions": "Quick Actions",
                "load_config": "Load Config from File",
                "lang_switch": "Language (语言)",
                "sender_settings": "1. Sender Settings (Recommended: Tencent Exmail)",
                "smtp_server": "SMTP Server:",
                "smtp_port": "SMTP Port:",
                "sender_email": "Sender Email:",
                "email_auth_code": "Authorization Code:",
                "data_files": "2. Data and File Selection",
                "select_source_file": "Select Source Data File (Excel):",
                "browse": "Browse...",
                "select_split_field": "Select Field for Splitting:",
                "english_config": "English Processing - Select values to process in English:",
                "select_mapping_file": "Select Email Mapping File (Excel):",
                "select_save_location": "Select Save Location for Split Files:",
                "email_content": "3. Email Content Configuration",
                "subject_prefix": "Email Subject (Prefix):",
                "use_filename_as_prefix": "Use source filename as prefix",
                "cc_recipients": "CC (separate multiple with ';'):",
                "chinese_prefix": "Chinese Email Body Prefix:",
                "chinese_suffix": "Chinese Email Body Suffix:",
                "english_prefix": "English Email Body Prefix:",
                "english_suffix": "English Email Body Suffix:",
                "execution_progress": "4. Execution and Progress",
                "execute_button": "Execute All (with Smart CN/EN Handling)",
                "success_title": "Success",
                "config_loaded_msg": "Configuration has been successfully loaded from the file.",
                "error_title": "Error",
                "config_load_error_msg": "Could not read or parse the configuration file.\nPlease ensure it has two columns (Key, Value) and is correctly formatted.\n\nDetails: {}",
                "file_read_error_msg": "Could not read file headers or map columns: {}",
                "all_tasks_complete_title": "Complete",
                "all_tasks_complete_msg": "All emails have been sent successfully! Includes smart CN/EN handling and statistical analysis.",
                "execution_error_title": "Execution Error",
                "execution_error_msg": "Task interrupted. Please check the log window for error messages.\n\nDetails: {}"
            }
        }
        self.current_lang = 'zh' # Default language

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
        self.english_processing_values = set()
        self.split_field_values = []

        # --- Sheet name mappings ---
        self.SHEET_NAMES = {
            'chinese': {
                "details": "details", "2x_stern": "累计2次严厉警告员工明细", "3x_stern": "累计3次及以上严厉警告员工明细", "branch_risk": "网点员工风险预警"
            },
            'english': {
                "details": "details", "2x_stern": "Staff_with_2x_Stern_Reminders", "3x_stern": "Staff_with_3x+_Stern_Reminders", "branch_risk": "Branch_Staff_Risk_Alert"
            }
        }
        
        # --- Field mappings for English sheets ---
        self.FIELD_MAPPINGS = {
            'chinese': {
                "满2次严厉警告员工人数": "满2次严厉警告员工人数", "超3次及以上严厉警告员工人数": "超3次及以上严厉警告员工人数"
            },
            'english': {
                "满2次严厉警告员工人数": "EmployeeCount_with_2x_Stern_Reminders", "超3次及以上严厉警告员工人数": "EmployeeCount_with_3x+_Stern_Reminders"
            }
        }

        self.use_filename_as_subject_var = tk.BooleanVar(value=False)

        self.setup_ui()
        self.update_ui_language()

    def setup_ui(self):
        self.geometry("1200x900")

        # --- Main canvas with a scrollbar ---
        self.canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.bind_mousewheel()

        # --- Root layout frame (inside the scrollable area) ---
        root_frame = ttk.Frame(self.scrollable_frame, padding="10")
        root_frame.pack(fill=tk.BOTH, expand=True)
        root_frame.columnconfigure(0, weight=1) # Main content column expands
        root_frame.columnconfigure(1, weight=0) # Sidebar column has fixed width

        self.create_widgets(root_frame)

    def create_widgets(self, root_frame):
        # --- Left Column: Main Configuration ---
        left_column_frame = ttk.Frame(root_frame)
        left_column_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        # --- Right Column: Quick Actions ---
        right_column_frame = ttk.Frame(root_frame)
        right_column_frame.grid(row=0, column=1, sticky="ns")

        # --- Language Switcher ---
        lang_frame = ttk.LabelFrame(right_column_frame, padding="10")
        self.lang_frame_label = lang_frame
        lang_frame.pack(fill=tk.X, pady=5)
        
        lang_buttons_frame = ttk.Frame(lang_frame)
        lang_buttons_frame.pack()
        ttk.Button(lang_buttons_frame, text="中文", command=lambda: self.switch_language('zh'), width=8).pack(side=tk.LEFT, padx=5)
        ttk.Button(lang_buttons_frame, text="English", command=lambda: self.switch_language('en'), width=8).pack(side=tk.LEFT, padx=5)
        
        # --- Config Loader ---
        config_frame = ttk.LabelFrame(right_column_frame, padding="10")
        self.config_frame_label = config_frame
        config_frame.pack(fill=tk.X, pady=5)
        
        self.load_config_button = ttk.Button(config_frame, command=self.load_configuration_file)
        self.load_config_button.pack(pady=10, fill=tk.X, ipady=4)

        # --- 1. Sender Settings ---
        self.sender_frame = ttk.LabelFrame(left_column_frame, padding="10")
        self.sender_frame.pack(fill=tk.X, pady=5)
        
        self.smtp_server_label = ttk.Label(self.sender_frame)
        self.smtp_server_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.smtp_server_var = tk.StringVar(value="smtp.exmail.qq.com")
        ttk.Entry(self.sender_frame, textvariable=self.smtp_server_var, width=30).grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.smtp_port_label = ttk.Label(self.sender_frame)
        self.smtp_port_label.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.smtp_port_var = tk.StringVar(value="465")
        ttk.Entry(self.sender_frame, textvariable=self.smtp_port_var, width=10).grid(row=1, column=3, padx=5, pady=5)

        self.sender_email_label = ttk.Label(self.sender_frame)
        self.sender_email_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.sender_email_var = tk.StringVar()
        ttk.Entry(self.sender_frame, textvariable=self.sender_email_var, width=30).grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        self.password_label = ttk.Label(self.sender_frame)
        self.password_label.grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.password_var = tk.StringVar()
        ttk.Entry(self.sender_frame, textvariable=self.password_var, show="*", width=20).grid(row=2, column=3, padx=5, pady=5, sticky="ew")
        
        # --- 2. Data and File Selection ---
        self.data_frame = ttk.LabelFrame(left_column_frame, padding="10")
        self.data_frame.pack(fill=tk.X, pady=5)
        self.data_frame.columnconfigure(1, weight=1)

        self.source_file_label = ttk.Label(self.data_frame)
        self.source_file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.source_file_var = tk.StringVar()
        ttk.Entry(self.data_frame, textvariable=self.source_file_var).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.browse_button1 = ttk.Button(self.data_frame, command=self.select_source_file)
        self.browse_button1.grid(row=0, column=2, padx=5, pady=5)
        
        self.split_column_label = ttk.Label(self.data_frame)
        self.split_column_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.split_column_var = tk.StringVar()
        self.split_column_combo = ttk.Combobox(self.data_frame, textvariable=self.split_column_var, state="readonly")
        self.split_column_combo.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.split_column_combo.bind('<<ComboboxSelected>>', self.on_split_column_selected)

        self.english_config_label = ttk.Label(self.data_frame)
        self.english_config_label.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        
        self.english_values_frame = ttk.Frame(self.data_frame)
        self.english_values_frame.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        self.english_checkboxes, self.english_checkbox_widgets = {}, {}

        self.mapping_file_label = ttk.Label(self.data_frame)
        self.mapping_file_label.grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.mapping_file_var = tk.StringVar()
        ttk.Entry(self.data_frame, textvariable=self.mapping_file_var).grid(row=4, column=1, padx=5, pady=5, sticky="ew")
        self.browse_button2 = ttk.Button(self.data_frame, command=self.select_mapping_file)
        self.browse_button2.grid(row=4, column=2, padx=5, pady=5)

        self.save_dir_label = ttk.Label(self.data_frame)
        self.save_dir_label.grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.save_dir_var = tk.StringVar()
        ttk.Entry(self.data_frame, textvariable=self.save_dir_var).grid(row=5, column=1, padx=5, pady=5, sticky="ew")
        self.browse_button3 = ttk.Button(self.data_frame, command=self.select_save_directory)
        self.browse_button3.grid(row=5, column=2, padx=5, pady=5)

        # --- 3. Email Content ---
        self.content_frame = ttk.LabelFrame(left_column_frame, padding="10")
        self.content_frame.pack(fill=tk.X, pady=5)
        self.content_frame.columnconfigure(1, weight=1)

        self.subject_label = ttk.Label(self.content_frame)
        self.subject_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.subject_var = tk.StringVar()
        self.subject_entry = ttk.Entry(self.content_frame, textvariable=self.subject_var)
        self.subject_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        self.subject_checkbox = ttk.Checkbutton(self.content_frame, variable=self.use_filename_as_subject_var, command=self.toggle_subject_entry_state)
        self.subject_checkbox.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        self.cc_label = ttk.Label(self.content_frame)
        self.cc_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.cc_var = tk.StringVar()
        ttk.Entry(self.content_frame, textvariable=self.cc_var).grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        self.cn_prefix_label = ttk.Label(self.content_frame)
        self.cn_prefix_label.grid(row=2, column=0, padx=5, pady=5, sticky="nw")
        self.chinese_prefix_text = tk.Text(self.content_frame, height=4)
        self.chinese_prefix_text.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        self.cn_suffix_label = ttk.Label(self.content_frame)
        self.cn_suffix_label.grid(row=3, column=0, padx=5, pady=5, sticky="nw")
        self.chinese_suffix_text = tk.Text(self.content_frame, height=4)
        self.chinese_suffix_text.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        self.en_prefix_label = ttk.Label(self.content_frame)
        self.en_prefix_label.grid(row=4, column=0, padx=5, pady=5, sticky="nw")
        self.english_prefix_text = tk.Text(self.content_frame, height=4)
        self.english_prefix_text.grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        self.en_suffix_label = ttk.Label(self.content_frame)
        self.en_suffix_label.grid(row=5, column=0, padx=5, pady=5, sticky="nw")
        self.english_suffix_text = tk.Text(self.content_frame, height=4)
        self.english_suffix_text.grid(row=5, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        # --- 4. Execution and Progress ---
        self.exec_frame = ttk.LabelFrame(left_column_frame, padding="10")
        self.exec_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.start_button = ttk.Button(self.exec_frame, command=self.start_sending_thread)
        self.start_button.pack(pady=10, ipady=4)

        self.progress = ttk.Progressbar(self.exec_frame, orient="horizontal", mode="determinate")
        self.progress.pack(pady=5, fill=tk.X, padx=5)
        
        # --- Log Area with its own Scrollbar ---
        log_frame = ttk.Frame(self.exec_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.status_log = tk.Text(log_frame, height=12, state="disabled")
        self.status_log.grid(row=0, column=0, sticky="nsew")

        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.status_log.yview)
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.status_log['yscrollcommand'] = log_scrollbar.set

    def switch_language(self, lang):
        self.current_lang = lang
        self.update_ui_language()

    def update_ui_language(self):
        lang_dict = self.LANG[self.current_lang]
        self.title(lang_dict["title"])
        
        # Quick Actions
        self.lang_frame_label.config(text=lang_dict["lang_switch"])
        self.config_frame_label.config(text=lang_dict["quick_actions"])
        self.load_config_button.config(text=lang_dict["load_config"])

        # Sender Settings
        self.sender_frame.config(text=lang_dict["sender_settings"])
        self.smtp_server_label.config(text=lang_dict["smtp_server"])
        self.smtp_port_label.config(text=lang_dict["smtp_port"])
        self.sender_email_label.config(text=lang_dict["sender_email"])
        self.password_label.config(text=lang_dict["email_auth_code"])

        # Data and Files
        self.data_frame.config(text=lang_dict["data_files"])
        self.source_file_label.config(text=lang_dict["select_source_file"])
        self.browse_button1.config(text=lang_dict["browse"])
        self.browse_button2.config(text=lang_dict["browse"])
        self.browse_button3.config(text=lang_dict["browse"])
        self.split_column_label.config(text=lang_dict["select_split_field"])
        self.english_config_label.config(text=lang_dict["english_config"])
        self.mapping_file_label.config(text=lang_dict["select_mapping_file"])
        self.save_dir_label.config(text=lang_dict["select_save_location"])

        # Email Content
        self.content_frame.config(text=lang_dict["email_content"])
        self.subject_label.config(text=lang_dict["subject_prefix"])
        self.subject_checkbox.config(text=lang_dict["use_filename_as_prefix"])
        self.cc_label.config(text=lang_dict["cc_recipients"])
        self.cn_prefix_label.config(text=lang_dict["chinese_prefix"])
        self.cn_suffix_label.config(text=lang_dict["chinese_suffix"])
        self.en_prefix_label.config(text=lang_dict["english_prefix"])
        self.en_suffix_label.config(text=lang_dict["english_suffix"])

        # Execution
        self.exec_frame.config(text=lang_dict["execution_progress"])
        self.start_button.config(text=lang_dict["execute_button"])
        
    def toggle_subject_entry_state(self):
        if self.use_filename_as_subject_var.get():
            self.subject_entry.config(state="disabled")
            self.subject_var.set("")
        else:
            self.subject_entry.config(state="normal")

    def bind_mousewheel(self):
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
            title="Select Configuration File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if not filepath:
            return

        try:
            df = pd.read_csv(filepath, header=None) if filepath.endswith('.csv') else pd.read_excel(filepath, header=None)
            config_dict = pd.Series(df.iloc[:, 1].values, index=df.iloc[:, 0]).dropna().to_dict()

            UI_VARS_MAP = {
                "smtp_server": self.smtp_server_var, "smtp_port": self.smtp_port_var,
                "sender_email": self.sender_email_var, "password": self.password_var,
                "subject_prefix": self.subject_var, "cc_recipients": self.cc_var,
                "chinese_prefix": self.chinese_prefix_text, "chinese_suffix": self.chinese_suffix_text,
                "english_prefix": self.english_prefix_text, "english_suffix": self.english_suffix_text
            }

            for key, value in config_dict.items():
                widget = UI_VARS_MAP.get(key)
                if isinstance(widget, tk.StringVar):
                    widget.set(str(value))
                elif isinstance(widget, tk.Text):
                    widget.delete("1.0", tk.END)
                    widget.insert("1.0", str(value))
            
            self.log("Configuration loaded successfully.")
            messagebox.showinfo(self.LANG[self.current_lang]["success_title"], self.LANG[self.current_lang]["config_loaded_msg"])

        except Exception as e:
            error_msg = self.LANG[self.current_lang]["config_load_error_msg"].format(e)
            self.log(f"Error: Failed to load config file: {e}")
            messagebox.showerror(self.LANG[self.current_lang]["error_title"], error_msg)

    def _get_column_mappings(self, df_columns):
        self.log("Starting smart column name mapping...")
        self.COLUMN_MAP = {}
        df_columns_lower = {col.lower().strip(): col for col in df_columns}

        for key, possible_names in self.KEY_COLS.items():
            found = False
            for name in possible_names:
                for lower_col, original_col in df_columns_lower.items():
                    if name in lower_col:
                        self.COLUMN_MAP[key] = original_col
                        self.log(f"  ✓ Mapped successfully: '{key}' -> '{original_col}'")
                        found = True
                        break
                if found:
                    break
        
        for crit_col in self.CRITICAL_COLS:
            if crit_col not in self.COLUMN_MAP:
                raise ValueError(f"Critical column missing! The program could not find a column representing '{crit_col}'. Please ensure the file contains a header with one of the keywords: {self.KEY_COLS[crit_col]}")

        self.log("Smart column name mapping complete.")
        return self.COLUMN_MAP

    def select_source_file(self):
        filepath = filedialog.askopenfilename(title="Select Source Data File", filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if filepath:
            self.source_file_var.set(filepath)
            try:
                df = pd.read_csv(filepath) if filepath.endswith('.csv') else pd.read_excel(filepath)
                
                self.split_column_combo['values'] = list(df.columns)
                self.log(f"Successfully loaded source file: {os.path.basename(filepath)}")
                self.log("Please select the field for splitting from the dropdown menu.")
                
                self.source_df = df
                self._get_column_mappings(df.columns)
            except Exception as e:
                error_msg = self.LANG[self.current_lang]["file_read_error_msg"].format(e)
                messagebox.showerror(self.LANG[self.current_lang]["error_title"], error_msg)
                self.log(f"Error: Could not read or map file {os.path.basename(filepath)}. Details: {e}")

    def on_split_column_selected(self, event=None):
        if not hasattr(self, 'source_df') or self.source_df is None: return
        
        split_column = self.split_column_var.get()
        if not split_column: return
        
        try:
            unique_values = sorted(self.source_df[split_column].dropna().unique())
            self.split_field_values = [str(val) for val in unique_values]
            
            for widget in self.english_checkbox_widgets.values(): widget.destroy()
            self.english_checkboxes.clear()
            self.english_checkbox_widgets.clear()
            
            self.log(f"Detected {len(self.split_field_values)} unique values in split field '{split_column}'")
            
            max_cols = 4
            for i, value in enumerate(self.split_field_values):
                row, col = i // max_cols, i % max_cols
                var = tk.BooleanVar()
                self.english_checkboxes[value] = var
                checkbox = ttk.Checkbutton(self.english_values_frame, text=f"{value}", variable=var, command=self.update_english_processing_values)
                checkbox.grid(row=row, column=col, padx=10, pady=2, sticky="w")
                self.english_checkbox_widgets[value] = checkbox
            
            self.log("You can now select values that require full English processing (multiple or none).")
            
        except Exception as e:
            self.log(f"Error processing split field: {e}")

    def update_english_processing_values(self):
        self.english_processing_values = {value for value, var in self.english_checkboxes.items() if var.get()}
        
        if self.english_processing_values:
            self.log(f"Selected for English processing: {', '.join(sorted(self.english_processing_values))}")
        else:
            self.log("No values are currently selected for English processing.")

    def select_mapping_file(self):
        filepath = filedialog.askopenfilename(title="Select Email Mapping File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.mapping_file_var.set(filepath)
            self.log(f"Selected email mapping file: {os.path.basename(filepath)}")

    def select_save_directory(self):
        dirpath = filedialog.askdirectory(title="Select Save Location")
        if dirpath:
            self.save_dir_var.set(dirpath)
            self.log(f"Split files will be saved to: {dirpath}")
            
    def start_sending_thread(self):
        self.start_button.config(state="disabled")
        self.status_log.config(state="normal")
        self.status_log.delete(1.0, tk.END)
        self.status_log.config(state="disabled")
        
        self.update_english_processing_values()
        if self.english_processing_values:
            self.log(f"Final confirmation for English processing: {', '.join(sorted(self.english_processing_values))}")
        else:
            self.log("No values selected for English processing; will use default Chinese mode.")
        
        processing_thread = threading.Thread(target=self.process_and_send_emails)
        processing_thread.daemon = True 
        processing_thread.start()

    def preprocess_data(self, df):
        self.log("Starting data preprocessing and standardization...")
        df = df.copy() 

        warning_type_col = self.COLUMN_MAP.get('warning_type')
        if warning_type_col:
            self.log(f"Standardizing '{warning_type_col}' column...")
            df[warning_type_col] = df[warning_type_col].astype(str)

            def standardize_warning_type(value):
                value = str(value).strip()
                if re.search('[\u4e00-\u9fff]', value):
                    if "严厉" in value: return "严厉警告"
                    elif "口述" in value: return "口述警告"
                    return value 
                else:
                    if "stern" in value.lower(): return "Stern Reminder"
                    elif "verbal" in value.lower(): return "Verbal Warning"
                    return value
            
            df[warning_type_col] = df[warning_type_col].apply(standardize_warning_type)
            self.log("  - 'Warning Type' field values unified based on original language (CN/EN).")

        sending_status_col = self.COLUMN_MAP.get('sending_status')
        if sending_status_col:
            self.log(f"Standardizing '{sending_status_col}' column...")
            df[sending_status_col] = df[sending_status_col].astype(str)
            df.loc[df[sending_status_col].str.contains("待发送", na=False), sending_status_col] = "已发送"
            df.loc[df[sending_status_col].str.contains("Pending", na=False, case=False), sending_status_col] = "Has been sent"
            self.log("  - 'Sending Status' field values unified based on keywords.")

        self.log("Data preprocessing and standardization complete.")
        return df

    def is_english_processing_required(self, area_name):
        return str(area_name).strip() in self.english_processing_values

    def generate_sheet_names(self, is_english=False):
        return self.SHEET_NAMES['english' if is_english else 'chinese']

    def count_warnings_per_employee(self, df):
        id_col, warning_type_col = self.COLUMN_MAP.get('id'), self.COLUMN_MAP.get('warning_type')
        if not id_col or not warning_type_col:
            self.log("Error: Cannot count warnings as employee ID or warning type column is not mapped.")
            return {}

        mapping_for_stats = {"严厉警告": "Stern Reminder", "口述警告": "Verbal Warning"}
        temp_warning_types = df[warning_type_col].replace(mapping_for_stats)
        
        warning_counts_df = df.groupby([id_col, temp_warning_types]).size().unstack(fill_value=0)
        
        result = {}
        for emp_id in df[id_col].unique():
            counts = warning_counts_df.loc[emp_id] if emp_id in warning_counts_df.index else {}
            result[emp_id] = {"Stern Reminder": counts.get("Stern Reminder", 0), "Verbal Warning": counts.get("Verbal Warning", 0)}
        return result

    def generate_warning_analysis_sheets(self, df_split, all_employees_warning_counts, is_english=False):
        sheets_data = {}
        sheet_names = self.generate_sheet_names(is_english)
        field_mappings = self.FIELD_MAPPINGS['english' if is_english else 'chinese']
        
        id_col, branch_col = self.COLUMN_MAP.get('id'), self.COLUMN_MAP.get('branch')
        base_columns_keys = ['area', 'district', 'branch', 'ops', 'position', 'id', 'name', 'status', 'employment_type']
        available_base_columns = [self.COLUMN_MAP[key] for key in base_columns_keys if key in self.COLUMN_MAP]
        
        unique_employees_in_split = df_split.drop_duplicates(subset=[id_col])

        two_stern_employees, three_plus_stern_employees = [], []
        for _, emp_data in unique_employees_in_split.iterrows():
            emp_id = emp_data[id_col]
            counts = all_employees_warning_counts.get(emp_id, {"Stern Reminder": 0, "Verbal Warning": 0})
            
            row_data = emp_data[available_base_columns].to_dict()
            if is_english:
                row_data["Stern Reminder"], row_data["Verbal Warning"] = counts["Stern Reminder"], counts["Verbal Warning"]
            else:
                row_data["严厉警告 Stern Reminder"], row_data["口述警告 Verbal Warning"] = counts["Stern Reminder"], counts["Verbal Warning"]
            
            if counts["Stern Reminder"] == 2:
                two_stern_employees.append(row_data)
            elif counts["Stern Reminder"] >= 3:
                three_plus_stern_employees.append(row_data)
        
        if two_stern_employees:
            df_2x = pd.DataFrame(two_stern_employees).sort_values(list(two_stern_employees[0].keys())[-2], ascending=False)
            sheets_data[sheet_names["2x_stern"]] = df_2x
            self.log(f"Analysis generated '{sheet_names['2x_stern']}': {len(two_stern_employees)} employees")
        
        if three_plus_stern_employees:
            df_3x = pd.DataFrame(three_plus_stern_employees).sort_values(list(three_plus_stern_employees[0].keys())[-2], ascending=False)
            sheets_data[sheet_names["3x_stern"]] = df_3x
            self.log(f"Analysis generated '{sheet_names['3x_stern']}': {len(three_plus_stern_employees)} employees")
        
        if branch_col:
            branch_risk_data = []
            for branch in df_split[branch_col].dropna().unique():
                branch_employees_df = df_split[df_split[branch_col] == branch]
                branch_employee_ids = branch_employees_df[id_col].unique()
                count_2x = sum(1 for emp_id in branch_employee_ids if all_employees_warning_counts.get(emp_id, {}).get("Stern Reminder", 0) == 2)
                count_3x_plus = sum(1 for emp_id in branch_employee_ids if all_employees_warning_counts.get(emp_id, {}).get("Stern Reminder", 0) >= 3)
                if count_2x > 0 or count_3x_plus > 0:
                    branch_info = branch_employees_df.iloc[0]
                    risk_row = {
                        self.COLUMN_MAP.get('area', 'Area'): branch_info.get(self.COLUMN_MAP.get('area')),
                        self.COLUMN_MAP.get('district', 'District'): branch_info.get(self.COLUMN_MAP.get('district')),
                        branch_col: branch, self.COLUMN_MAP.get('ops', 'OPS'): branch_info.get(self.COLUMN_MAP.get('ops')),
                        field_mappings["满2次严厉警告员工人数"]: count_2x, field_mappings["超3次及以上严厉警告员工人数"]: count_3x_plus
                    }
                    branch_risk_data.append(risk_row)
            
            if branch_risk_data:
                branch_risk_df = pd.DataFrame(branch_risk_data)
                branch_risk_df["_total_risk"] = branch_risk_df[field_mappings["满2次严厉警告员工人数"]].fillna(0) + branch_risk_df[field_mappings["超3次及以上严厉警告员工人数"]].fillna(0)
                branch_risk_df = branch_risk_df.sort_values("_total_risk", ascending=False).drop("_total_risk", axis=1)
                sheets_data[sheet_names["branch_risk"]] = branch_risk_df
                self.log(f"Analysis generated '{sheet_names['branch_risk']}': {len(branch_risk_df)} branches")
        
        return sheets_data

    def generate_statistics_summary(self, df_split, all_employees_warning_counts, is_english=False):
        id_col, branch_col, status_col = self.COLUMN_MAP.get('id'), self.COLUMN_MAP.get('branch'), self.COLUMN_MAP.get('status')
        unique_employees = df_split.drop_duplicates(subset=[id_col])
        
        count_2x_total, count_3x_plus_total = 0, 0
        status_2x_count, status_3x_plus_count = {}, {}
        
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
            for branch in df_split[branch_col].dropna().unique():
                details = {'total': 0, '2x_count': 0, '3x_plus_count': 0, '2x_status_breakdown': {}, '3x_plus_status_breakdown': {}}
                branch_employees = df_split[df_split[branch_col] == branch].drop_duplicates(subset=[id_col])
                for _, emp_data in branch_employees.iterrows():
                    emp_id, status = emp_data[id_col], emp_data.get(status_col, "未知" if not is_english else "Unknown")
                    counts = all_employees_warning_counts.get(emp_id, {})
                    if counts.get("Stern Reminder", 0) == 2:
                        details['2x_count'] += 1
                        details['2x_status_breakdown'][status] = details['2x_status_breakdown'].get(status, 0) + 1
                    elif counts.get("Stern Reminder", 0) >= 3:
                        details['3x_plus_count'] += 1
                        details['3x_plus_status_breakdown'][status] = details['3x_plus_status_breakdown'].get(status, 0) + 1
                details['total'] = details['2x_count'] + details['3x_plus_count']
                if details['total'] > 0: branch_risk_details[branch] = details
        
        top_5_branches = sorted(branch_risk_details.items(), key=lambda x: x[1]['total'], reverse=True)[:5]
        
        if is_english:
            return f"""=== WARNING STATISTICS SUMMARY ===

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
===============================""".strip()
        else:
            return f"""=== 警告统计总结 ===

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
====================""".strip()

    def _format_status_breakdown(self, status_count, is_english=False):
        if not status_count: return "- 无数据" if not is_english else "- No data"
        return "\n".join([f"- {status}: {count}{'人' if not is_english else ''}" for status, count in status_count.items()])

    def _format_status_breakdown_inline(self, breakdown_dict, is_english=False):
        if not breakdown_dict: return ""
        items = [f"{status}: {count}{'人' if not is_english else ''}" for status, count in breakdown_dict.items()]
        return ", ".join(items)

    def _format_top_branches_detailed(self, top_branches, is_english=False):
        if not top_branches: return "- 无数据" if not is_english else "- No data"
        result = []
        for i, (branch, details) in enumerate(top_branches, 1):
            if is_english:
                branch_str = f"{i}. {branch} - Total Risk: {details['total']}"
                if details['3x_plus_count'] > 0:
                    branch_str += f"\n   - Employees with 3+ Stern Reminders: {details['3x_plus_count']}; Status: {self._format_status_breakdown_inline(details['3x_plus_status_breakdown'], is_english)}"
                if details['2x_count'] > 0:
                    branch_str += f"\n   - Employees with 2 Stern Reminders: {details['2x_count']}; Status: {self._format_status_breakdown_inline(details['2x_status_breakdown'], is_english)}"
            else:
                branch_str = f"{i}. {branch} - 总风险人数: {details['total']}人"
                if details['3x_plus_count'] > 0:
                    branch_str += f"\n   - 累计3次及以上严厉警告员工: {details['3x_plus_count']}人；状态: {self._format_status_breakdown_inline(details['3x_plus_status_breakdown'], is_english)}"
                if details['2x_count'] > 0:
                    branch_str += f"\n   - 累计2次严厉警告员工: {details['2x_count']}人；状态: {self._format_status_breakdown_inline(details['2x_status_breakdown'], is_english)}"
            result.append(branch_str)
        return "\n".join(result)
    
    def create_multi_sheet_excel(self, df_split, file_path, area_name, all_employees_warning_counts):
        is_english = self.is_english_processing_required(area_name)
        sheet_names = self.generate_sheet_names(is_english)
        analysis_sheets = self.generate_warning_analysis_sheets(df_split, all_employees_warning_counts, is_english)
        
        ordered_sheets = []
        if sheet_names["branch_risk"] in analysis_sheets: ordered_sheets.append((sheet_names["branch_risk"], analysis_sheets[sheet_names["branch_risk"]]))
        for sheet_key in ["3x_stern", "2x_stern"]:
            if sheet_names[sheet_key] in analysis_sheets: ordered_sheets.append((sheet_names[sheet_key], analysis_sheets[sheet_names[sheet_key]]))
        ordered_sheets.append((sheet_names["details"], df_split))
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, sheet_df in ordered_sheets:
                sheet_df_copy = sheet_df.copy()
                stat_cols_to_clear = [
                    "严厉警告 Stern Reminder", "口述警告 Verbal Warning", "Stern Reminder", "Verbal Warning",
                    self.FIELD_MAPPINGS['chinese']["满2次严厉警告员工人数"], self.FIELD_MAPPINGS['english']["满2次严厉警告员工人数"],
                    self.FIELD_MAPPINGS['chinese']["超3次及以上严厉警告员工人数"], self.FIELD_MAPPINGS['english']["超3次及以上严厉警告员工人数"]
                ]
                if sheet_name != sheet_names["details"]:
                    for col in stat_cols_to_clear:
                        if col in sheet_df_copy.columns: sheet_df_copy[col] = sheet_df_copy[col].replace(0, None)
                sheet_df_copy.to_excel(writer, sheet_name=sheet_name, index=False)
        self.log(f"Multi-sheet Excel file generated: {os.path.basename(file_path)} ({'English Mode' if is_english else 'Chinese Mode'})")

    def generate_email_content(self, area_name, df_split, all_employees_warning_counts):
        is_english = self.is_english_processing_required(area_name)
        
        prefix = self.english_prefix_text.get("1.0", tk.END).strip() if is_english else self.chinese_prefix_text.get("1.0", tk.END).strip()
        suffix = self.english_suffix_text.get("1.0", tk.END).strip() if is_english else self.chinese_suffix_text.get("1.0", tk.END).strip()
        summary = self.generate_statistics_summary(df_split, all_employees_warning_counts, is_english)
        
        return f"{prefix}\n\n{summary}\n\n{suffix}".strip()

    def process_and_send_emails(self):
        try:
            self.log("Starting task, checking parameters...")
            
            subject_prefix = ""
            if self.use_filename_as_subject_var.get():
                source_file_path = self.source_file_var.get()
                if not source_file_path: raise ValueError("Please select a source data file before using its name as the subject.")
                subject_prefix = os.path.splitext(os.path.basename(source_file_path))[0]
                self.log(f"Email subject prefix set to source filename: '{subject_prefix}'")
            else:
                subject_prefix = self.subject_var.get()

            params = {
                "source_file": self.source_file_var.get(), "split_column": self.split_column_var.get(),
                "mapping_file": self.mapping_file_var.get(), "save_dir": self.save_dir_var.get(),
                "smtp_server": self.smtp_server_var.get(), "smtp_port": self.smtp_port_var.get(),
                "sender_email": self.sender_email_var.get(), "password": self.password_var.get(),
                "subject_prefix": subject_prefix, "cc_recipients": [cc.strip() for cc in self.cc_var.get().split(';') if cc.strip()]
            }
            
            for field in ["source_file", "split_column", "mapping_file", "save_dir", "smtp_server", "smtp_port", "sender_email", "password", "subject_prefix"]:
                if not params[field]:
                    if field == "subject_prefix": raise ValueError("Email subject prefix cannot be empty! Please fill it or check the 'use filename' option.")
                    raise ValueError(f"Required field '{field}' is empty! Please check your configuration.")
            
            cn_prefix, cn_suffix = self.chinese_prefix_text.get("1.0", tk.END).strip(), self.chinese_suffix_text.get("1.0", tk.END).strip()
            en_prefix, en_suffix = self.english_prefix_text.get("1.0", tk.END).strip(), self.english_suffix_text.get("1.0", tk.END).strip()
            if not (cn_prefix or cn_suffix) and not (en_prefix or en_suffix):
                raise ValueError("At least one email template (Chinese or English prefix/suffix) must be filled.")
            
            self.log("Parameter validation passed.")

            self.log("Reading source data file...")
            df_source = pd.read_csv(params["source_file"]) if params["source_file"].endswith('.csv') else pd.read_excel(params["source_file"])
            self.log(f"Source data file contains {len(df_source)} rows.")
            
            self._get_column_mappings(df_source.columns)
            df_source_preprocessed = self.preprocess_data(df_source)
            
            self.log("Performing global count of warnings for all employees...")
            all_employees_warning_counts = self.count_warnings_per_employee(df_source_preprocessed)
            self.log(f"Global count complete. Analyzed {len(all_employees_warning_counts)} unique employees.")
            
            self.log("Reading email mapping file...")
            df_mapping = pd.read_excel(params["mapping_file"])
            mapping_dict = pd.Series(df_mapping.iloc[:, 1].values, index=df_mapping.iloc[:, 0]).to_dict()
            self.log("Email mapping loaded successfully.")

            split_values = df_source_preprocessed[params["split_column"]].dropna().unique()
            total_tasks = len(split_values)
            self.progress['maximum'] = total_tasks
            self.log(f"Detected {total_tasks} unique split values to process.")

            for i, value in enumerate(split_values):
                self.log("-" * 60)
                self.log(f"Processing: [{value}] ({i+1}/{total_tasks})")

                df_split = df_source_preprocessed[df_source_preprocessed[params["split_column"]] == value].copy()
                self.log(f"Split data contains {len(df_split)} rows.")
                
                is_english_processing = self.is_english_processing_required(value)
                processing_mode = "English Mode" if is_english_processing else "Chinese Mode"
                self.log(f"Current processing mode: {processing_mode}")
                
                attachment_filename = f"{params['subject_prefix']}_{value}.xlsx"
                attachment_path = os.path.join(params["save_dir"], attachment_filename)
                
                self.create_multi_sheet_excel(df_split, attachment_path, value, all_employees_warning_counts)

                recipient_email = mapping_dict.get(value)
                if not recipient_email:
                    self.log(f"Warning: No email found for '{value}' in the mapping file. Skipping this item.")
                    self.progress['value'] = i + 1
                    continue
                self.log(f"Found recipient: {recipient_email}")

                email_body = self.generate_email_content(value, df_split, all_employees_warning_counts)
                
                msg = MIMEMultipart()
                msg['From'], msg['To'], msg['Subject'] = params["sender_email"], recipient_email, f"{params['subject_prefix']}_{value}"
                if params["cc_recipients"]: msg['Cc'] = ";".join(params["cc_recipients"])
                msg.attach(MIMEText(email_body, 'plain', 'utf-8'))

                try:
                    with open(attachment_path, "rb") as f:
                        part = MIMEApplication(f.read(), Name=attachment_filename)
                    part['Content-Disposition'] = f'attachment; filename="{attachment_filename}"'
                    msg.attach(part)
                    self.log(f"Attached multi-sheet Excel file: {attachment_filename}")
                except Exception as attach_error:
                    self.log(f"Error attaching file: {attach_error}")
                    continue

                all_recipients = [recipient_email] + params["cc_recipients"]
                self.log(f"Connecting to SMTP server: {params['smtp_server']}:{params['smtp_port']}...")
                try:
                    with smtplib.SMTP_SSL(params["smtp_server"], int(params["smtp_port"])) as server:
                        server.login(params["sender_email"], params["password"])
                        self.log("Server login successful, sending email...")
                        server.sendmail(params["sender_email"], all_recipients, msg.as_string())
                        cc_info = f"CC: {';'.join(params['cc_recipients'])}" if params["cc_recipients"] else "No CC"
                        self.log(f"Email sent successfully to: {recipient_email} ({cc_info}) - Mode: {processing_mode}")
                except Exception as email_error:
                    self.log(f"Email sending failed: {email_error}")
                    continue

                self.progress['value'] = i + 1

            self.log("-" * 60)
            self.log("All tasks completed!")
            messagebox.showinfo(self.LANG[self.current_lang]["all_tasks_complete_title"], self.LANG[self.current_lang]["all_tasks_complete_msg"])

        except Exception as e:
            error_msg = self.LANG[self.current_lang]["execution_error_msg"].format(e)
            self.log(f"A fatal error occurred: {e}")
            messagebox.showerror(self.LANG[self.current_lang]["execution_error_title"], error_msg)
        finally:
            self.start_button.config(state="normal")

if __name__ == "__main__":
    app = EmailSenderApp()
    app.mainloop()