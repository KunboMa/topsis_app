import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np

def calculate_topsis(df, selected_columns, direction_list):
    df_vars = df[selected_columns].astype(float)

    # 标准化
    df_std = df_vars.copy()
    for i, col in enumerate(selected_columns):
        if direction_list[i] == 1:
            df_std[col] = (df_vars[col] - df_vars[col].min()) / (df_vars[col].max() - df_vars[col].min())
        elif direction_list[i] == -1:
            df_std[col] = (df_vars[col].max() - df_vars[col]) / (df_vars[col].max() - df_vars[col].min())
        else:
            raise ValueError("方向必须是 1（极大化）或 -1（极小化）")

    # 熵权法
    X = df_std.values
    n, m = X.shape
    P = X / X.sum(axis=0)
    epsilon = 1e-12
    k = 1 / np.log(n)
    E = -k * np.sum(P * np.log(P + epsilon), axis=0)
    d = 1 - E
    w = d / d.sum()

    # 加权规范化矩阵
    weighted_matrix = df_std * w

    # 正负理想解
    Z_plus = weighted_matrix.max(axis=0)
    Z_neg = weighted_matrix.min(axis=0)

    # 距离计算
    D_plus = np.sqrt(((weighted_matrix - Z_plus) ** 2).sum(axis=1))
    D_minus = np.sqrt(((weighted_matrix - Z_neg) ** 2).sum(axis=1))

    # 综合评价值 CR
    CR = D_minus / (D_plus + D_minus)

    result = pd.DataFrame({
        "FCIL_CDE": df["FCIL_CDE"].values,
        "TOPSIS得分": CR
    }).sort_values(by="TOPSIS得分", ascending=False).reset_index(drop=True)
    
    return result


class TopsisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("TOPSIS 变量选择界面")
        self.file_path = ""
        self.df = None
        self.vars = []
        self.direction_vars = {}
        # 计算按钮
        self.run_btn = tk.Button(root, text="✅ 开始计算并导出结果", command=self.run_topsis)
        self.run_btn.pack(pady=10)


        # 上传按钮
        self.upload_btn = tk.Button(root, text="📁 上传 Excel 文件", command=self.load_excel)
        self.upload_btn.pack(pady=10)

        # 变量选择区域
        self.var_frame = tk.LabelFrame(root, text="选择用于计算的变量（并设定方向）")
        self.var_frame.pack(padx=10, pady=10, fill="both")

    def load_excel(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if not self.file_path:
            return

        try:
            self.df = pd.read_excel(self.file_path)
            print("✅ 成功读取文件:", self.file_path)  # ← 添加这一行
            print(self.df.head())  # ← 再加这一行：打印内容确认结构
            self.show_variable_options()
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取 Excel 文件：{str(e)}")

    def show_variable_options(self):
        for widget in self.var_frame.winfo_children():
            widget.destroy()

        numeric_cols = self.df.select_dtypes(include=[np.number]).columns.tolist()

        self.vars = []
        self.direction_vars = {}

        for col in numeric_cols:
            row = tk.Frame(self.var_frame)
            row.pack(fill="x", padx=5, pady=2)

            var_check = tk.IntVar()
            chk = tk.Checkbutton(row, text=col, variable=var_check, width=40, anchor="w")
            chk.pack(side="left")
            self.vars.append((col, var_check))

            dir_var = tk.StringVar(value="极大化")
            dir_menu = tk.OptionMenu(row, dir_var, "极大化", "极小化")
            dir_menu.pack(side="left")
            self.direction_vars[col] = dir_var
            
    def run_topsis(self):
        selected_columns = [col for col, var_check in self.vars if var_check.get() == 1]
        if not selected_columns:
            messagebox.showwarning("未选择变量", "请至少选择一个变量进行计算")
            return

    # 提取方向（极大 or 极小）
        direction_list = []
        for col in selected_columns:
            direction = self.direction_vars[col].get()
            direction_list.append(1 if direction == "极大化" else -1)
            
        try:
            result_df = calculate_topsis(self.df, selected_columns, direction_list)

        # 输出结果 Excel
            output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            title="保存计算结果为..."
            )
            if output_path:
                result_df.to_excel(output_path, index=False)
                messagebox.showinfo("完成", f"TOPSIS 计算完成，结果已保存至：\n{output_path}")
        except Exception as e:
            messagebox.showerror("计算错误", f"发生错误：{str(e)}")


# 启动界面
if __name__ == "__main__":
    root = tk.Tk()
    app = TopsisGUI(root)
    root.mainloop()