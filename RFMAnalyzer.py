import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, ttk
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

        # 滚动区域容器
        self.scroll_frame = tk.Frame(master)
        self.scroll_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.scroll_frame)
        self.scrollbar = tk.Scrollbar(self.scroll_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.inner_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # 错误消息框
        self.error_msg = tk.Text(self.inner_frame, height=3, width=50)
        self.error_msg.insert(tk.END, "这里是报错信息展示区域")
        self.error_msg.pack()

        # 动态评分组管理
        self.score_frames = {'R': self.create_score_group(self.inner_frame, "R评分标准"),
                             'F': self.create_score_group(self.inner_frame, "F评分标准"),
                             'M': self.create_score_group(self.inner_frame, "M评分标准")}

        self.load_param_btn = tk.Button(self.inner_frame, text="加载参数文件", command=self.load_params)
        self.load_param_btn.pack()

        self.compare_label = tk.Label(self.inner_frame, text="对比值 (格式: R,F,M):")
        self.compare_label.pack(pady=10)
        self.compare_entry = tk.Entry(self.inner_frame, width=50)
        self.compare_entry.pack(pady=10)

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
        self.update_group_fields(group_frame, type_var)

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

    def load_params(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if not file_path:
            return

        try:
            self.clear_score_groups()  # 清空已有评分组
            self.parse_txt_params(file_path)
        except Exception as e:
            self.show_error(f"加载参数文件时发生错误: {str(e)}")

    def parse_txt_params(self, file_path):
        with open(file_path, "r", encoding="utf-8") as file:
            for line in file:
                line = line.strip()
                if not line or ":" not in line:
                    continue

                category, params = line.split(":")
                if category not in self.score_frames:
                    continue

                params = params.split(",")
                if len(params) < 3:
                    self.show_error("参数格式错误，请检查文件内容。")
                    return

                score_type = params[0]
                score_val = params[-1]

                if score_type == "区间":
                    if "-" not in params[1]:
                        self.show_error("区间格式错误，应为 min-max 格式。")
                        return
                    min_val, max_val = map(int, params[1].split("-"))
                elif score_type == "大于":
                    min_val, max_val = int(params[1]), None
                elif score_type == "小于":
                    min_val, max_val = None, int(params[1])
                else:
                    self.show_error(f"不支持的评分类型: {score_type}")
                    return

                self.add_param_to_frame(self.score_frames[category], score_type, min_val, max_val, int(score_val))

    def clear_score_groups(self):
        for frame in self.score_frames.values():
            for child in frame.winfo_children():
                child.destroy()

    def add_param_to_frame(self, frame, score_type, min_val, max_val, score_val):
        group_frame = tk.Frame(frame)
        group_frame.pack(pady=5)

        type_var = tk.StringVar(value=score_type)
        type_menu = ttk.OptionMenu(group_frame, type_var, score_type, "区间", "大于", "小于",
                                   command=lambda x: self.update_group_fields(group_frame, type_var))
        type_menu.pack(side=tk.LEFT)

        min_entry = tk.Entry(group_frame, width=5)
        if min_val is not None:
            min_entry.insert(0, str(min_val))
        min_entry.pack(side=tk.LEFT)

        spacer = tk.Label(group_frame, text="到")
        spacer.pack(side=tk.LEFT)

        max_entry = tk.Entry(group_frame, width=5)
        if max_val is not None:
            max_entry.insert(0, str(max_val))
        max_entry.pack(side=tk.LEFT)

        score_label = tk.Label(group_frame, text="Score:")
        score_label.pack(side=tk.LEFT)

        score_entry = tk.Entry(group_frame, width=5)
        score_entry.insert(0, str(score_val))
        score_entry.pack(side=tk.LEFT)

        group_frame.entries = {"type": type_var, "min": min_entry, "max": max_entry, "score": score_entry}
        self.update_group_fields(group_frame, type_var)

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
            except (ValueError, TypeError):  # 捕获非整数或空值错误
                self.show_error("无效的输入: 请确保所有评分字段为整数。")
                continue

        # 确保 score_dict 中的所有值都是 int
        score_dict = {k: int(v) for k, v in score_dict.items()}

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
        print(message)  # 打印错误信息到控制台

    @staticmethod
    def apply_custom_scores(value, score_dict, reverse=False):
        try:
            value = int(value)  # 确保 value 是整数
            for key, score in score_dict.items():
                if isinstance(key, tuple):
                    if len(key) == 2 and isinstance(key[0], int) and isinstance(key[1], int) and key[0] <= value <= key[1]:
                        return score if not reverse else max(score_dict.values()) - score + 1
                elif key[0] == '>' and value > key[1]:
                    return score
                elif key[0] == '<' and value < key[1]:
                    return score
        except Exception as e:
            print(f"Error in apply_custom_scores: {e}")
        return 0

    def calculate_rfm(self, data, r_scores, f_scores, m_scores):
        try:
            # 计算 R、F、M 的得分
            data['R_Score'] = data['recency'].apply(lambda x: self.apply_custom_scores(x, r_scores, reverse=True))
            data['F_Score'] = data['frequency'].apply(lambda x: self.apply_custom_scores(x, f_scores))
            data['M_Score'] = data['monetary'].apply(lambda x: self.apply_custom_scores(x, m_scores))

            rfm = data[['role_id', 'R_Score', 'F_Score', 'M_Score']].copy()
            rfm['RFM_Score'] = rfm['R_Score'] + rfm['F_Score'] + rfm['M_Score']
            
            return rfm
        except Exception as e:
            self.show_error(f"计算 RFM 分数时发生错误: {str(e)}")
            raise

    def segment_users(self, rfm_data, avg_recency, avg_frequency, avg_monetary):
        try:
            rfm_data = rfm_data.copy()
            rfm_data['Segment'] = rfm_data.apply(self.define_segment, axis=1, args=(avg_recency, avg_frequency, avg_monetary))

            # 根据选择的图表类型执行相应逻辑
            selected_option = self.segment_var.get()
            if selected_option == "RFM用户组别频次图":
                self.generate_frequency_chart(rfm_data)
        except Exception as e:
            self.show_error(f"用户分层时发生错误: {str(e)}")
            raise

    def define_segment(self, row, avg_recency, avg_frequency, avg_monetary):
        try:
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
        except Exception as e:
            self.show_error(f"定义用户分层时发生错误: {str(e)}")
            raise

    def generate_frequency_chart(self, rfm_data):
        try:
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
                percentage = segment_percentages.iloc[i]
                ax.text(i, count, f'{count} ({percentage:.2f}%)', ha='center', va='bottom')

            plt.xticks(rotation=45)
            plt.tight_layout()

            # 设置字体以支持中文
            plt.rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
            plt.rcParams['axes.unicode_minus'] = False  # 正常显示负号

            plt.show()
        except Exception as e:
            self.show_error(f"生成频次图时发生错误: {str(e)}")
            raise

if __name__ == "__main__":
    root = tk.Tk()
    app = RFMAnalyzer(root)
    root.mainloop()