import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

class RFMAnalyzer:
    def __init__(self, master):
        self.master = master
        self.master.title("RFM 用户分层")
        
        self.label = tk.Label(master, text="导入 CSV/XLSX 文件:")
        self.label.pack()
        
        self.upload_btn = tk.Button(master, text="上传文件", command=self.upload_file)
        self.upload_btn.pack()
        
        self.segment_var = tk.StringVar(master)
        self.segment_options = [
            "RFM用户组别频次图",
            "RFM描述分析"
        ]
        self.segment_menu = ttk.OptionMenu(master, self.segment_var, self.segment_options[0], *self.segment_options)
        self.segment_menu.pack()
        
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
            elif self.file_path.endswith(".xlsx"):
                data = pd.read_excel(self.file_path)
            else:
                raise ValueError("不支持的文件类型")
            
            data = data[['role_id', 'recency', 'frequency', 'monetary']]
            
            # 执行 RFM 分析
            rfm_data = self.calculate_rfm(data)
            self.segment_users(rfm_data)
        
        except Exception as e:
            self.error_msg.delete(1.0, tk.END)
            self.error_msg.insert(tk.END, f"发生错误: {str(e)}")
    
    @staticmethod
    def min_max_normalize(column):
        return (column - column.min()) / (column.max() - column.min())

    @staticmethod
    def min_max_normalize_inverse(column):
        return 1 - (column - column.min()) / (column.max() - column.min())

    @staticmethod
    def add_noise(column):
        noise = np.random.uniform(0, 1e-9, size=column.shape)
        return column + noise
    
    def calculate_rfm(self, data):
        # 减少鲸鱼玩家付费金额和付费频次的影响，使用对数转换来降低极端值影响
        data[['monetary', 'frequency']] = np.log1p(data[['monetary', 'frequency']])
        
        # 对RFM进行归一化处理,其中R值需要进行逆向归一化
        data['Normalized_Recency'] = self.min_max_normalize_inverse(data['recency'])
        data['Normalized_Frequency'] = self.min_max_normalize(data['frequency'])
        data['Normalized_Monetary'] = self.min_max_normalize(data['monetary'])
        
        # 增加微小的随机噪声以处理重复值
        data['Normalized_Recency'] = self.add_noise(data['Normalized_Recency'])
        data['Normalized_Frequency'] = self.add_noise(data['Normalized_Frequency'])
        data['Normalized_Monetary'] = self.add_noise(data['Normalized_Monetary'])
        
        # 再次检查并处理仍然存在的重复值
        while data.duplicated(['Normalized_Recency', 'Normalized_Frequency', 'Normalized_Monetary']).any():
            data['Normalized_Recency'] = self.add_noise(data['Normalized_Recency'])
            data['Normalized_Frequency'] = self.add_noise(data['Normalized_Frequency'])
            data['Normalized_Monetary'] = self.add_noise(data['Normalized_Monetary'])
        
        rfm = pd.DataFrame()
        
        # 计算 RFM 每个维度的得分
        rfm['R_Score'] = pd.qcut(data['Normalized_Recency'], 4, labels=[1, 2, 3, 4]).astype(int)
        rfm['F_Score'] = pd.qcut(data['Normalized_Frequency'], 4, labels=[1, 2, 3, 4]).astype(int)
        rfm['M_Score'] = pd.qcut(data['Normalized_Monetary'], 4, labels=[1, 2, 3, 4]).astype(int)
        
        # 计算 RFM 总分
        rfm['RFM_Score'] = rfm['R_Score'].astype(int) + rfm['F_Score'].astype(int) + rfm['M_Score'].astype(int)
        
        return rfm

    def segment_users(self, rfm_data):
        avg_recency = rfm_data['R_Score'].mean()
        avg_frequency = rfm_data['F_Score'].mean()
        avg_monetary = rfm_data['M_Score'].mean()
        
        # 定义用户分层函数
        def segment_customer(row):
            if row['R_Score'] > avg_recency:
                if row['F_Score'] > avg_frequency and row['M_Score'] > avg_monetary:
                    return '高价值客户'
                elif row['F_Score'] > avg_frequency and row['M_Score'] <= avg_monetary:
                    return '重点发展客户'
                elif row['F_Score'] <= avg_frequency and row['M_Score'] > avg_monetary:
                    return '重点保持客户'
                else:
                    return '重点挽留客户'
            else:
                if row['F_Score'] > avg_frequency and row['M_Score'] > avg_monetary:
                    return '一般价值客户'
                elif row['F_Score'] > avg_frequency and row['M_Score'] <= avg_monetary:
                    return '一般发展客户'
                elif row['F_Score'] <= avg_frequency and row['M_Score'] > avg_monetary:
                    return '一般保持客户'
                else:
                    return '一般挽留客户'
        
        # 应用细分函数
        rfm_data['Segment'] = rfm_data.apply(segment_customer, axis=1)
        
        # 根据选择的图表类型执行相应逻辑
        selected_option = self.segment_var.get()
        if selected_option == "RFM用户组别频次图":
            self.generate_frequency_chart(rfm_data)

    def generate_frequency_chart(self, rfm_data):
        # 计算每个细分组的频次和占比
        segment_counts = rfm_data['Segment'].value_counts()
        total_count = len(rfm_data)
        segment_percentages = (segment_counts / total_count) * 100

        # 创建结果数据框
        summary_df = pd.DataFrame({
            '用户分类': segment_counts.index,
            '频率': segment_counts.values,
            '占比': segment_percentages.round(2).astype(str) + '%'
        })

        # 打印汇总信息
        print(summary_df)

        # 绘制频次图
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.bar(segment_counts.index, segment_counts.values, color='skyblue', alpha=0.7)
        ax.set_ylabel('频率')
        ax.set_title('RFM 用户组别频次图')
        
        # 添加频次和占比标签
        for i, count in enumerate(segment_counts):
            percentage = segment_percentages[i]
            ax.text(i, count, f'{count} ({percentage:.2f}%)', ha='center', va='bottom')

        plt.xticks(rotation=45)
        plt.tight_layout()
        
        # 设置字体以支持中文
        plt.rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
        plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号
        
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = RFMAnalyzer(root)
    root.mainloop()