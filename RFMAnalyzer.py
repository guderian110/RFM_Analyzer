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

        # 添加自定义评分标准的输入框
        self.r_label = tk.Label(master, text="R评分标准 (格式: min1-max1:score1,min2-max2:score2,...):")
        self.r_label.pack()
        self.r_entry = tk.Entry(master, width=50)
        self.r_entry.pack()

        self.f_label = tk.Label(master, text="F评分标准 (格式: min1-max1:score1,min2-max2:score2,...):")
        self.f_label.pack()
        self.f_entry = tk.Entry(master, width=50)
        self.f_entry.pack()

        self.m_label = tk.Label(master, text="M评分标准 (格式: min1-max1:score1,min2-max2:score2,...):")
        self.m_label.pack()
        self.m_entry = tk.Entry(master, width=50)
        self.m_entry.pack()

        self.compare_label = tk.Label(master, text="对比值 (格式: R,F,M):")
        self.compare_label.pack()
        self.compare_entry = tk.Entry(master, width=50)
        self.compare_entry.pack()

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
    def parse_custom_scores(entry):
        scores = entry.split(',')
        score_dict = {}
        for score in scores:
            range_, value = score.split(':')
            min_, max_ = map(int, range_.split('-'))
            score_dict[(min_, max_)] = int(value)
        return score_dict

    @staticmethod
    def apply_custom_scores(value, score_dict):
        for (min_, max_), score in score_dict.items():
            if min_ <= value <= max_:
                return score
        return 0

    def calculate_rfm(self, data):
        # 获取用户自定义评分标准
        r_scores = self.parse_custom_scores(self.r_entry.get())
        f_scores = self.parse_custom_scores(self.f_entry.get())
        m_scores = self.parse_custom_scores(self.m_entry.get())

        # 计算 R、F、M 的得分
        data['R_Score'] = data['recency'].apply(lambda x: self.apply_custom_scores(x, r_scores))
        data['F_Score'] = data['frequency'].apply(lambda x: self.apply_custom_scores(x, f_scores))
        data['M_Score'] = data['monetary'].apply(lambda x: self.apply_custom_scores(x, m_scores))

        rfm = data[['role_id', 'R_Score', 'F_Score', 'M_Score']]

        # 计算 RFM 总分
        rfm['RFM_Score'] = rfm['R_Score'] + rfm['F_Score'] + rfm['M_Score']
        
        return rfm

    def segment_users(self, rfm_data):
        # 获取用户输入的对比值
        compare_values = self.compare_entry.get()
        if compare_values:
            avg_recency, avg_frequency, avg_monetary = map(int, compare_values.split(','))
        else:
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