import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, ttk
import os
from tkinter import messagebox


class RFMAnalyzer:
    def __init__(self, master):
        self.master = master
        self.master.title("RFM 用户分层分析系统")
        self.master.geometry("1000x800")
        self.master.configure(bg="#f0f3f5")

        # 样式配置
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.configure_styles()
        
        # 主容器
        self.main_frame = ttk.Frame(master)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 顶部操作区
        self.create_top_panel()
        
        # 分析结果显示区
        self.create_result_panel()
        
        # 参数设置区
        self.create_parameter_panel()

        # 初始化变量
        self.file_path = ""
        self.current_params = {"R": {}, "F": {}, "M": {}}

    def configure_styles(self):
        # 自定义样式
        self.style.configure("TButton", padding=6, font=("微软雅黑", 10))
        self.style.configure("TLabel", font=("微软雅黑", 10), background="#f0f3f5")
        self.style.configure("Header.TLabel", font=("微软雅黑", 12, "bold"))
        self.style.configure("Section.TFrame", background="#ffffff", relief=tk.RIDGE, borderwidth=2)
        self.style.configure("Error.TLabel", foreground="#dc3545", font=("微软雅黑", 9))
        self.style.map("Accent.TButton",
                      foreground=[("active", "#ffffff"), ("!disabled", "#ffffff")],
                      background=[("active", "#0062cc"), ("!disabled", "#007bff")])
        
    def create_top_panel(self):
        # 顶部操作面板
        top_frame = ttk.Frame(self.main_frame, style="Section.TFrame")
        top_frame.pack(fill=tk.X, pady=(0, 15))

        # 文件上传区
        upload_frame = ttk.Frame(top_frame)
        upload_frame.pack(pady=10, padx=10, fill=tk.X)
        
        ttk.Label(upload_frame, text="数据文件:", style="Header.TLabel").pack(side=tk.LEFT)
        self.upload_btn = ttk.Button(upload_frame, text="选择文件", style="Accent.TButton", command=self.upload_file)
        self.upload_btn.pack(side=tk.LEFT, padx=10)
        self.file_label = ttk.Label(upload_frame, text="未选择文件", foreground="#6c757d")
        self.file_label.pack(side=tk.LEFT)

        # 分析控制区
        ctrl_frame = ttk.Frame(top_frame)
        ctrl_frame.pack(pady=10, padx=10, fill=tk.X)
        
        ttk.Label(ctrl_frame, text="分析类型:").pack(side=tk.LEFT)
        self.segment_var = tk.StringVar()
        self.segment_menu = ttk.OptionMenu(ctrl_frame, self.segment_var, "RFM用户组别频次图", 
                                          "RFM用户组别频次图", "RFM描述分析")
        self.segment_menu.pack(side=tk.LEFT, padx=10)
        
        self.analyze_btn = ttk.Button(ctrl_frame, text="开始分析", style="Accent.TButton", command=self.analyze_data)
        self.analyze_btn.pack(side=tk.RIGHT)    
    def create_result_panel(self):
        # 结果展示面板
        result_frame = ttk.Frame(self.main_frame, style="Section.TFrame")
        result_frame.pack(fill=tk.BOTH, expand=True)

        # 错误提示
        self.error_label = ttk.Label(result_frame, text="", style="Error.TLabel")
        self.error_label.pack(pady=5)

        # 带滚动条的画布
        self.canvas = tk.Canvas(result_frame, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.param_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.param_frame, anchor="nw")

        self.param_frame.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))

    def create_parameter_panel(self):
        # 参数设置面板
        param_container = ttk.Frame(self.param_frame)
        param_container.pack(pady=10, padx=10, fill=tk.X)

        # 评分标准设置
        score_frame = ttk.LabelFrame(param_container, text="评分标准设置", padding=(10, 5))
        score_frame.pack(fill=tk.X, pady=5)
        
        self.score_frames = {
            'R': self.create_score_group(score_frame, "最近消费间隔（R）"),
            'F': self.create_score_group(score_frame, "消费频率（F）"),
            'M': self.create_score_group(score_frame, "消费金额（M）")
        }

        # 参数管理
        param_btn_frame = ttk.Frame(score_frame)
        param_btn_frame.pack(pady=10)
        
        ttk.Button(param_btn_frame, text="加载参数", command=self.load_params).pack(side=tk.LEFT, padx=5)
        # ttk.Button(param_btn_frame, text="保存参数", command=self.save_params).pack(side=tk.LEFT, padx=5)

        # 对比值设置
        compare_frame = ttk.Frame(param_container)
        compare_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(compare_frame, text="基准值设置（R,F,M）:").pack(side=tk.LEFT)
        self.compare_entry = ttk.Entry(compare_frame, width=20)
        self.compare_entry.pack(side=tk.LEFT, padx=10)
        ttk.Label(compare_frame, text="示例：30,5,1000", foreground="#6c757d").pack(side=tk.LEFT)

    def create_score_group(self, parent, title):
        frame = ttk.LabelFrame(parent, text=title, padding=(10, 5))
        frame.pack(fill=tk.X, pady=5, padx=5)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=5)
        
        ttk.Button(btn_frame, text="+ 添加条件", command=lambda: self.add_score_group(frame)).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="- 删除条件", command=lambda: self.remove_score_group(frame)).pack(side=tk.LEFT)
        
        return frame

        return group_frame
    def show_error(self, message):
        self.error_label.config(text=message)
        self.master.after(5000, lambda: self.error_label.config(text=""))  # 5秒后自动清除错误信息

    def upload_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("数据文件", "*.csv *.xlsx")],
            title="选择数据文件"
        )
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.show_error("")  # 清除旧错误信息
    def add_score_group(self, frame):
        group_frame = ttk.Frame(frame)
        group_frame.pack(pady=5)

        type_var = tk.StringVar(value="区间")
        type_menu = ttk.OptionMenu(group_frame, type_var, "区间", "区间", "大于", "小于",
                                    command=lambda x: self.update_group_fields(group_frame, type_var))
        type_menu.pack(side=tk.LEFT)

        min_entry = ttk.Entry(group_frame, width=5)
        min_entry.pack(side=tk.LEFT)

        spacer = ttk.Label(group_frame, text="到")
        spacer.pack(side=tk.LEFT)

        max_entry = ttk.Entry(group_frame, width=5)
        max_entry.pack(side=tk.LEFT)

        score_label = ttk.Label(group_frame, text="Score:")
        score_label.pack(side=tk.LEFT)

        score_entry = ttk.Entry(group_frame, width=5)
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

    def remove_score_group(self, parent_frame):
        # 获取所有子组件（排除按钮框架）
        children = [child for child in parent_frame.winfo_children() if not isinstance(child, ttk.Frame)]
        
        # 删除最后一个条件条目
        if len(children) > 0:
            children[-1].destroy()

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
        group_frame = ttk.Frame(frame)
        group_frame.pack(pady=5)

        type_var = tk.StringVar(value=score_type)
        type_menu = ttk.OptionMenu(group_frame, type_var, score_type, "区间", "大于", "小于",
                                   command=lambda x: self.update_group_fields(group_frame, type_var))
        type_menu.pack(side=tk.LEFT)

        min_entry = ttk.Entry(group_frame, width=5)
        if min_val is not None:
            min_entry.insert(0, str(min_val))
        min_entry.pack(side=tk.LEFT)

        spacer = ttk.Label(group_frame, text="到")
        spacer.pack(side=tk.LEFT)

        max_entry = ttk.Entry(group_frame, width=5)
        if max_val is not None:
            max_entry.insert(0, str(max_val))
        max_entry.pack(side=tk.LEFT)

        score_label = ttk.Label(group_frame, text="Score:")
        score_label.pack(side=tk.LEFT)

        score_entry = ttk.Entry(group_frame, width=5)
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

        return score_dict

    def analyze_data(self):
        if not self.file_path:
            self.show_error("请上传文件。")
            return

        try:
            r_scores = self.get_score_groups(self.score_frames['R'])
            f_scores = self.get_score_groups(self.score_frames['F'])
            m_scores = self.get_score_groups(self.score_frames['M'])

            if not (r_scores and f_scores and m_scores):
                self.show_error("R、F、M评分标准为必填项，请填写所有评分标准。")
                return

            data = self.load_data(self.file_path)
            data = data[['role_id', 'recency', 'frequency', 'monetary']]

            rfm_data = self.calculate_rfm(data, r_scores, f_scores, m_scores)

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

            fig, ax = plt.subplots(figsize=(10, 6))
            ax.bar(segment_counts.index, segment_counts.values, color='skyblue', alpha=0.7)
            ax.set_ylabel('频率')
            ax.set_title('RFM 用户组别频次图')

            for i, count in enumerate(segment_counts):
                percentage = segment_percentages.iloc[i]
                ax.text(i, count, f'{count} ({percentage:.2f}%)', ha='center', va='bottom')

            plt.xticks(rotation=45)
            plt.tight_layout()

            plt.rcParams['font.sans-serif'] = ['SimHei']
            plt.rcParams['axes.unicode_minus'] = False

            plt.show()
        except Exception as e:
            self.show_error(f"生成频次图时发生错误: {str(e)}")
            raise


if __name__ == "__main__":
    root = tk.Tk()
    app = RFMAnalyzer(root)
    root.mainloop()