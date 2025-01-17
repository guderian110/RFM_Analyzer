import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
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
        self.segment_options = ["RFM用户组别频次图", "RFM描述分析"]
        self.segment_menu = ttk.OptionMenu(master, self.segment_var, self.segment_options[0], *self.segment_options)
        self.segment_menu.pack()

        self.execute_btn = tk.Button(master, text="执行分析", command=self.analyze_data)
        self.execute_btn.pack()

        self.error_msg = tk.Text(master, height=3, width=30)
        self.error_msg.insert(tk.END,"这里是报错信息展示区域")
        self.error_msg.config(state=tk.DISABLED)
        self.error_msg.pack()

        # 添加动态评分组管理
        self.score_frames = {'R': self.create_score_group(master, "R评分标准"),
                             'F': self.create_score_group(master, "F评分标准"),
                             'M': self.create_score_group(master, "M评分标准")}

        self.compare_label = tk.Label(master, text="对比值 (格式: R,F,M):")
        self.compare_label.pack()
        self.compare_entry = tk.Entry(master, width=50)
        self.compare_entry.pack()

        self.file_path = ""

    def create_score_group(self, master, label_text):
        group_frame = tk.Frame(master)
        group_frame.pack(pady=5)

        label = tk.Label(group_frame, text=label_text)
        label.pack(side=tk.LEFT)

        btn_add = tk.Button(group_frame, text="+", command=lambda: self.add_score_group(group_frame))
        btn_add.pack(side=tk.LEFT)

        btn_remove = tk.Button(group_frame, text="-", command=lambda: self.remove_score_group(group_frame))
        btn_remove.pack(side=tk.LEFT)

        frame = tk.Frame(master)
        frame.pack()

        self.add_score_group(frame)  # 初始化评分组
        return frame

    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if file_path:
            self.file_path = file_path
            self.label.config(text=f"已上传: {os.path.basename(file_path)}")

    def add_score_group(self, frame):
        group_frame = tk.Frame(frame)
        group_frame.pack(pady=5)

        type_var = tk.StringVar(value="区间")
        type_menu = ttk.OptionMenu(group_frame, type_var, "区间", "区间", "大于", "小于",
                                    command=lambda x: self.update_group_fields(group_frame, type_var))
        type_menu.pack(side=tk.LEFT)

        min_entry = tk.Entry(group_frame, width=5)
        min_entry.pack(side=tk.LEFT)

        spacer = tk.Label(group_frame, text="到")
        spacer.pack(side=tk.LEFT)

        max_entry = tk.Entry(group_frame, width=5)
        max_entry.pack(side=tk.LEFT)

        score_label = tk.Label(group_frame, text="Score:")
        score_label.pack(side=tk.LEFT)

        score_entry = tk.Entry(group_frame, width=5)
        score_entry.pack(side=tk.LEFT)

        group_frame.entries = {"type": type_var, "min": min_entry, "max": max_entry, "score": score_entry}
        self.update_group_fields(group_frame, type_var)  # Initialize field states

    def update_group_fields(self, group_frame, type_var):
        entries = group_frame.entries
        if type_var.get() == "大于":
            entries["min"].configure(state="normal")
            entries["max"].delete(0, tk.END)
            entries["max"].configure(state="disabled")
        elif type_var.get() == "小于":
            entries["max"].configure(state="normal")
            entries["min"].delete(0, tk.END)
            entries["min"].configure(state="disabled")
        else:  # 区间
            entries["min"].configure(state="normal")
            entries["max"].configure(state="normal")

    def remove_score_group(self, frame):
        if len(frame.winfo_children()) > 0:
            frame.winfo_children()[-1].destroy()

    def get_score_groups(self, frame):
        score_dict = {}
        for group_frame in frame.winfo_children():
            entries = group_frame.entries
            try:
                score_type = entries["type"].get()
                score_val = int(entries["score"].get())
                if score_type == "大于":
                    min_val = int(entries["min"].get())
                    score_dict[(">", min_val)] = score_val
                elif score_type == "小于":
                    max_val = int(entries["max"].get())
                    score_dict[("<", max_val)] = score_val
                else:  # 区间
                    min_val = int(entries["min"].get())
                    max_val = int(entries["max"].get())
                    score_dict[(min_val, max_val)] = score_val
            except ValueError:
                continue
        return score_dict

    def analyze_data(self):
        if not self.file_path:
            self.show_error("请上传文件。")
            return

        try:
            r_scores = self.get_score_groups(self.score_frames['R'])
            f_scores = self.get_score_groups(self.score_frames['F'])
            m_scores = self.get_score_groups(self.score_frames['M'])

            # 检查评分标准是否填写
            if not (r_scores and f_scores and m_scores):
                self.show_error("R、F、M评分标准为必填项，请填写所有评分标准。")
                return

            data = self.load_data(self.file_path)
            data = data[['role_id', 'recency', 'frequency', 'monetary']]

            # 执行 RFM 分析
            rfm_data = self.calculate_rfm(data, r_scores, f_scores, m_scores)

            # 获取对比值，若未填写则使用分数的平均值
            compare_values = self.compare_entry.get()
            avg_recency = rfm_data['R_Score'].mean()
            avg_frequency = rfm_data['F_Score'].mean()
            avg_monetary = rfm_data['M_Score'].mean()

            if compare_values:
                try:
                    user_recency, user_frequency, user_monetary = map(int, compare_values.split(','))
                    avg_recency = user_recency
                    avg_frequency = user_frequency
                    avg_monetary = user_monetary
                except ValueError:
                    self.show_error("对比值格式不正确，请使用: R,F,M")
                    return

            self.segment_users(rfm_data, avg_recency, avg_frequency, avg_monetary)

        except Exception as e:
            self.show_error(f"发生错误: {str(e)}")

    def load_data(self, file_path):
        if file_path.endswith(".csv"):
            return pd.read_csv(file_path)
        elif file_path.endswith(".xlsx"):
            return pd.read_excel(file_path)
        else:
            raise ValueError("不支持的文件类型")

    def show_error(self, message):
        self.error_msg.delete(1.0, tk.END)
        self.error_msg.insert(tk.END, message)

    @staticmethod
    def apply_custom_scores(value, score_dict, reverse=False):
        for key, score in score_dict.items():
            if isinstance(key, tuple):
                if key[0] == '>' and value > key[1]:
                    return score
                elif key[0] == '<' and value < key[1]:
                    return score
                elif len(key) == 2 and key[0] <= value <= key[1]:
                    return score if not reverse else max(score_dict.values()) - score + 1
        return 0

    def calculate_rfm(self, data, r_scores, f_scores, m_scores):
        # 计算 R、F、M 的得分
        data['R_Score'] = data['recency'].apply(lambda x: self.apply_custom_scores(x, r_scores, reverse=True))
        data['F_Score'] = data['frequency'].apply(lambda x: self.apply_custom_scores(x, f_scores))
        data['M_Score'] = data['monetary'].apply(lambda x: self.apply_custom_scores(x, m_scores))

        rfm = data[['role_id', 'R_Score', 'F_Score', 'M_Score']]
        rfm['RFM_Score'] = rfm['R_Score'] + rfm['F_Score'] + rfm['M_Score']
        
        return rfm

    def segment_users(self, rfm_data, avg_recency, avg_frequency, avg_monetary):
        rfm_data['Segment'] = rfm_data.apply(self.define_segment, axis=1, args=(avg_recency, avg_frequency, avg_monetary))

        # 根据选择的图表类型执行相应逻辑
        selected_option = self.segment_var.get()
        if selected_option == "RFM用户组别频次图":
            self.generate_frequency_chart(rfm_data)

    def define_segment(self, row, avg_recency, avg_frequency, avg_monetary):
        if row['R_Score'] > avg_recency:
            if row['F_Score'] > avg_frequency and row['M_Score'] > avg_monetary:
                return '高价值客户'
            elif row['F_Score'] > avg_frequency:
                return '重点发展客户'
            elif row['M_Score'] > avg_monetary:
                return '重点保持客户'
            else:
                return '重点挽留客户'
        else:
            if row['F_Score'] > avg_frequency and row['M_Score'] > avg_monetary:
                return '一般价值客户'
            elif row['F_Score'] > avg_frequency:
                return '一般发展客户'
            elif row['M_Score'] > avg_monetary:
                return '一般保持客户'
            else:
                return '一般挽留客户'

    def generate_frequency_chart(self, rfm_data):
        segment_counts = rfm_data['Segment'].value_counts()
        total_count = len(rfm_data)
        segment_percentages = (segment_counts / total_count) * 100

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