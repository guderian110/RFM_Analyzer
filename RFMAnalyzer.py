import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import filedialog, messagebox
import os

class RFMAnalyzer:
    def __init__(self, master):
        self.master = master
        self.master.title("RFM 用户分层")
        
        self.label = tk.Label(master, text="导入 CSV/XLSX 文件:")
        self.label.pack()
        
        self.upload_btn = tk.Button(master, text="上传文件", command=self.upload_file)
        self.upload_btn.pack()
        
        self.execute_btn = tk.Button(master, text="执行分析", command=self.analyze_data)
        self.execute_btn.pack()
        
        self.error_msg = tk.Text(master, height=5, width=50)
        self.error_msg.pack()

def upload_file(self):
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
    self.file_path = file_path
    self.label.config(text=f"已上传: {os.path.basename(file_path)}")
    
def analyze_data(self):
    try:
        if self.file_path.endswith(".csv"):
            data = pd.read_csv(self.file_path)
            data = data[['role_id','recency','frequency','monetary']]
        elif self.file_path.endswith(".xlsx"):
            data = pd.read_excel(self.file_path)
            data = data[['role_id','recency','frequency','monetary']]
        else:
            raise ValueError("不支持的文件类型")
        
        # 执行 RFM 分析
        rfm_data = self.calculate_rfm(data)
        
    except Exception as e:
        self.error_msg.delete(1.0, tk.END)
        self.error_msg.insert(tk.END, f"发生错误: {str(e)}")
 
def min_max_normalize(column):
    return (column - column.min()) / (column.max() - column.min())
 
        
def calculate_rfm(self, data):
        # 计算 RFM 指标
    