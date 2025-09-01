import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os
from datetime import datetime

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("虚假类警告信数据处理工具 False Warning Letter Processor V1.0 ")
        self.root.geometry("720x420")

        # --- 变量 ---
        self.file_a_path = tk.StringVar()
        self.file_b_path = tk.StringVar()
        self.date_suffix = tk.StringVar(value=datetime.now().strftime('%m%d'))

        # --- UI 框架 ---
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 文件A选择 ---
        ttk.Label(main_frame, text="表格A (警告类型数据, 用于匹配虚假单号) / File A (Warning Type Datas for Matching)", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        frame_a = ttk.Frame(main_frame)
        frame_a.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 20))
        ttk.Entry(frame_a, textvariable=self.file_a_path, width=70, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        ttk.Button(frame_a, text="选择文件... / Select...", command=self.select_file_a).pack(side=tk.RIGHT, padx=(10, 0))

        # --- 文件B选择 ---
        ttk.Label(main_frame, text="表格B (待处理警告信数据) / File B (Warning Datas to be Processed)", font=("Microsoft YaHei UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(0, 5))
        frame_b = ttk.Frame(main_frame)
        frame_b.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 20))
        ttk.Entry(frame_b, textvariable=self.file_b_path, width=70, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        ttk.Button(frame_b, text="选择文件... / Select...", command=self.select_file_b).pack(side=tk.RIGHT, padx=(10, 0))

        # --- 文件名后缀 ---
        ttk.Label(main_frame, text="设置输出文件名后缀 (数据对应的日期) / Set Suffix for Output File (Date of Datas)", font=("Microsoft YaHei UI", 10, "bold")).grid(row=4, column=0, sticky="w", pady=(0, 5))
        ttk.Entry(main_frame, textvariable=self.date_suffix, width=20).grid(row=5, column=0, sticky="w")
        
        # --- 状态栏 ---
        self.status_label = ttk.Label(main_frame, text="准备就绪。请按顺序选择文件A和文件B。/ Ready. Please select File A and File B in order.", foreground="gray")
        self.status_label.grid(row=6, column=0, columnspan=3, sticky="w", pady=(20, 0))

        # --- 处理按钮 ---
        process_button = ttk.Button(main_frame, text="开始处理并生成报告 / Process and Generate Report", command=self.run_processing)
        process_button.grid(row=7, column=0, columnspan=3, pady=(20, 10), ipady=5, sticky="ew")

    def select_file_a(self):
        path = filedialog.askopenfilename(title="请选择表格A / Select File A", filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file_a_path.set(path)

            if not self.file_b_path.get():
                self.update_status("文件A已选择。现在请选择文件B。/ File A selected. Now, please select File B.")
            else:
                self.update_status("文件A和B均已选择。点击下方按钮开始处理。/ Both files selected. Click the button below to start.")

    def select_file_b(self):
        path = filedialog.askopenfilename(title="请选择表格B / Select File B", filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file_b_path.set(path)

            if not self.file_a_path.get():
                self.update_status("文件B已选择。现在请选择文件A。/ File B selected. Now, please select File A.")
            else:
                self.update_status("文件A和B均已选择。点击下方按钮开始处理。/ Both files selected. Click the button below to start.")

    def update_status(self, message, color="blue"):
        self.status_label.config(text=message, foreground=color)
        self.root.update_idletasks()

    def extract_bill_num(self, detail_string):
        """从违规详情字符串中提取虚假单号"""
        if not isinstance(detail_string, str):
            return None
        match = re.search(r"虚假单号:(.*?);", detail_string, re.IGNORECASE)
        return match.group(1).strip() if match else None

    def run_processing(self):
        # --- 输入验证 ---
        if not self.file_a_path.get() or not self.file_b_path.get():
            messagebox.showerror("错误 / Error", "请确保两个表格文件都已选择！\nPlease ensure both files are selected!")
            return

        try:
            self.update_status("正在读取文件... / Reading files...")
            # --- 1. 读取数据 ---
            df_a = pd.read_excel(self.file_a_path.get(), sheet_name='details_original_type')
            df_b = pd.read_excel(self.file_b_path.get(), sheet_name='details')

            # --- 2. 数据预处理 ---
            self.update_status("正在进行数据预处理... / Preprocessing data...")
            if '电话' in df_b.columns:
                df_b = df_b.drop(columns=['电话'])
            
            if '违规类型' in df_b.columns:
                df_b['违规类型'] = df_b['违规类型'].astype(str)
                df_b.loc[df_b['违规类型'].str.contains('虚假妥投', na=False), '违规类型'] = '虚假妥投'
                df_b.loc[df_b['违规类型'].str.contains('虚假标记', na=False), '违规类型'] = '虚假标记'
            else:
                raise KeyError("表格B中缺少关键字段【违规类型】/ Missing required column in File B: [违规类型]")
            
            # --- 3. 自动化匹配与初始化 ---
            self.update_status("正在匹配数据并初始化列... / Matching data and initializing columns...")
            df_b['辅助列-Waybill'] = df_b['违规详情'].apply(self.extract_bill_num)
            violation_map = pd.Series(df_a['Violation type'].values, index=df_a['false_bill_num']).to_dict()
            df_b['辅助1'] = pd.NA
            
            toutou_mask = df_b['违规类型'] == '虚假妥投'
            biaoji_mask = df_b['违规类型'] == '虚假标记'

            if toutou_mask.any():
                extracted_nums = df_b.loc[toutou_mask, '违规详情'].apply(self.extract_bill_num)
                df_b.loc[toutou_mask, '辅助1'] = extracted_nums.map(violation_map)

            df_b['警告信发出建议'] = ''
            df_b['发送方式'] = ''
            processed_mask = pd.Series([False] * len(df_b), index=df_b.index)
            
            self.update_status("正在应用核心处理逻辑... / Applying core processing logic...")

            # --- 核心处理逻辑 (V1.0 优化版) ---

            # --- 5. 处理【虚假妥投】记录 ---
            
            # 优先级 1: 处理离职人员 (最高优)
            mask = toutou_mask & df_b['在职状态'].isin(['离职', '待离职']) & ~processed_mask
            df_b.loc[mask, ['警告信发出建议', '发送方式']] = ['不发出NotSent', 'Bulk Send']
            processed_mask |= mask

            # 优先级 2: 根据 '处理意见' == '员工申诉，建议采纳' 进行判断
            mask_base = toutou_mask & (df_b['处理意见'] == '员工申诉，建议采纳') & ~processed_mask
            # V4.0: 整合所有“不发出”的关键词，提高效率和可维护性
            do_not_send_keywords = 'pod valid|cancelled|non-false|no warning|not send|not sent|no issue of warning'
            mask_ok = mask_base & df_b['处理备注'].str.contains(do_not_send_keywords, case=False, na=False)
            df_b.loc[mask_ok, ['警告信发出建议', '发送方式']] = ['不发出NotSent', 'Bulk Send']
            processed_mask |= mask_ok
            # 对于建议采纳但无明确不发出理由的，转为人工复核
            mask_recheck = mask_base & ~processed_mask
            df_b.loc[mask_recheck, ['警告信发出建议', '发送方式']] = ['Manual Recheck', 'Single Send']
            processed_mask |= mask_recheck

            # 优先级 3: 根据 '处理意见' == '员工申诉，理由不充分' 进行判断
            mask_base = toutou_mask & (df_b['处理意见'] == '员工申诉，理由不充分') & ~processed_mask
            mask_verbal = mask_base & df_b['处理备注'].str.contains('verbal', case=False, na=False)
            df_b.loc[mask_verbal, ['警告信发出建议', '发送方式']] = ['口述Verbal', 'Bulk Send']
            processed_mask |= mask_verbal
            mask_recheck = mask_base & ~processed_mask
            df_b.loc[mask_recheck, ['警告信发出建议', '发送方式']] = ['Manual Recheck', 'Single Send']
            processed_mask |= mask_recheck

            # 优先级 4: 根据 '处理意见' == '员工未申诉，或态度不好' 进行判断
            mask_base = toutou_mask & (df_b['处理意见'] == '员工未申诉，或态度不好') & ~processed_mask
            mask_stern = mask_base & df_b['处理备注'].str.contains('stern', case=False, na=False)
            
            # 细分stern下的情况
            mask_stern_verbal = mask_stern & (df_b['辅助1'] == '口述Verbal')
            df_b.loc[mask_stern_verbal, ['警告信发出建议', '发送方式']] = ['口述Verbal', 'Single Send']
            processed_mask |= mask_stern_verbal
            
            mask_stern_stern = mask_stern & (df_b['辅助1'] == '严厉Stern')
            df_b.loc[mask_stern_stern, ['警告信发出建议', '发送方式']] = ['严厉Stern-Manual Recheck', 'Bulk Send-Manual Recheck']
            processed_mask |= mask_stern_stern
            
            # 其他未申诉/态度不好的情况
            mask_recheck = mask_base & ~processed_mask
            df_b.loc[mask_recheck, ['警告信发出建议', '发送方式']] = ['Manual Recheck', 'Single Send']
            processed_mask |= mask_recheck

            # V4.0: 新增“兜底”规则，确保所有“虚假妥投”记录都有处理建议
            mask_fallback = toutou_mask & ~processed_mask
            df_b.loc[mask_fallback, ['警告信发出建议', '发送方式']] = ['Manual Recheck', 'Single Send']
            processed_mask |= mask_fallback

            # --- 6. 处理【虚假标记】记录 (逻辑结构优化) ---
            
            # 规则 1: 申诉采纳且明确不发警告
            mask = biaoji_mask & (df_b['处理意见'] == '员工申诉，建议采纳') & df_b['处理备注'].str.contains('No Warning', case=False, na=False) & ~processed_mask
            df_b.loc[mask, ['警告信发出建议', '发送方式']] = ['不发出NotSent', 'Bulk Send']
            processed_mask |= mask

            # V2.0: 合并相似规则，结构更清晰
            # 规则 2 & 3: 未申诉或理由不充分的情况
            mask_base = biaoji_mask & df_b['处理意见'].isin(['员工未申诉，或态度不好', '员工申诉，理由不充分']) & ~processed_mask
            
            mask_stern = mask_base & df_b['处理备注'].str.contains('Stern', case=False, na=False)
            df_b.loc[mask_stern, ['警告信发出建议', '发送方式']] = ['严厉Stern', 'Bulk Send']
            processed_mask |= mask_stern
            
            mask_verbal = mask_base & df_b['处理备注'].str.contains('Verbal', case=False, na=False)
            df_b.loc[mask_verbal, ['警告信发出建议', '发送方式']] = ['口述Verbal', 'Bulk Send']
            processed_mask |= mask_verbal

            # 规则 4: 其他所有情况 (“兜底”规则)
            mask_fallback = biaoji_mask & ~processed_mask
            df_b.loc[mask_fallback, ['警告信发出建议', '发送方式']] = ['Manual Recheck', 'Single Send']
            processed_mask |= mask_fallback
            
            self.update_status("处理完成，请选择保存位置。/ Processing complete, please select a save location.")
            # --- 7. 保存结果 ---
            file_name = f"虚假类警告信确认_{self.date_suffix.get()}.xlsx"
            save_path = filedialog.asksaveasfilename(
                initialfile=file_name,
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )

            if save_path:
                df_b.to_excel(save_path, index=False)
                self.update_status(f"文件已成功保存! / File saved successfully!", color="green")
                messagebox.showinfo("成功 / Success", f"处理完成！文件已保存至：\n{save_path}\n\nProcessing Complete! File saved to:\n{save_path}")
            else:
                self.update_status("用户取消了保存操作。/ Save operation cancelled by user.", color="orange")

        except FileNotFoundError:
            messagebox.showerror("错误 / Error", "文件未找到，请检查路径是否正确。\nFile not found, please check the file path.")
            self.update_status("操作失败：文件未找到。/ Operation failed: File not found.", color="red")
        except KeyError as e:
            messagebox.showerror("错误 / Error", f"Excel文件中缺少必要的列名或工作表: {e}\n请确认文件A包含'details_original_type'工作表, 文件B包含'details'工作表, 且所有必需列均存在。")
            self.update_status(f"操作失败：缺少列或工作表 {e}。/ Operation failed: Missing column or sheet {e}.", color="red")
        except Exception as e:
            messagebox.showerror("发生未知错误 / Unknown Error", str(e))
            self.update_status(f"操作失败：{e} / Operation failed: {e}", color="red")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
