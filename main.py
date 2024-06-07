import gurobipy as gp
from gurobipy import GRB
import io
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook

class CaptureOutput(list):
    """Class to capture stdout and stderr."""
    def __enter__(self):
        self._stdout = sys.stdout
        self._stderr = sys.stderr
        sys.stdout = self._stringio = io.StringIO()
        sys.stderr = self._stringio
        return self

    def __exit__(self, *args):
        self.extend(self._stringio.getvalue().splitlines())
        sys.stdout = self._stdout
        sys.stderr = self._stderr

def execute_model_from_file(model_file: str) -> tuple[gp.Model, list]:
    # 從文件讀取模型
    model = gp.read(model_file)

    # 捕獲 Gurobi 的日誌輸出
    with CaptureOutput() as output:
        # 求解模型
        model.optimize()
        model.printStats()

    log_output = output._stringio.getvalue().splitlines()
    
    # 顯示捕獲的日誌輸出到控制台
    for line in log_output:
        print(line)

    return model, log_output

def save_to_excel(model: gp.Model, log_output: list, filename: str):
    # 準備變數和值數據
    data = {
        'Variable': [],
        'Value': []
    }

    for v in model.getVars():
        data['Variable'].append(v.VarName)
        data['Value'].append(v.X)

    df = pd.DataFrame(data)

    # 準備附加信息（求解時間、迭代次數、目標函數值等）
    additional_info = {
        'Info': [
            'Objective Value', 
            'Solve Time (seconds)', 
            'Iterations', 
            'Node Count'
        ],
        'Value': [
            model.ObjVal, 
            model.Runtime, 
            model.IterCount, 
            model.NodeCount
        ]
    }

    df_info = pd.DataFrame(additional_info)

    # 捕獲的日誌輸出
    log_data = {
        'Log': log_output
    }
    df_log = pd.DataFrame(log_data)

    # 創建或加載 Excel 文件並寫入數據
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Solution', index=False)
        df_info.to_excel(writer, sheet_name='Info', index=False)
        df_log.to_excel(writer, sheet_name='Log', index=False)

def main():
    # 建立主窗口
    root = tk.Tk()
    root.title("Gurobi Model Runner")
    # 禁用視窗拉伸
    root.resizable(False, False)

    # 文件選擇函數
    def select_file():
        file_path = filedialog.askopenfilename(filetypes=[("Model Files", "*.lp *.mps *.json")])
        if file_path:
            model_file.set(file_path)

    # 執行模型求解
    def solve_model():
        try:
            model_file_path = model_file.get()
            if not model_file_path:
                messagebox.showwarning("Warning", "請選擇模型文件")
                return
            
            global main_model, log_output
            main_model, log_output = execute_model_from_file(model_file_path)
            
            if main_model.status == GRB.OPTIMAL:
                result = f"Optimal solution: {main_model.ObjVal}\nSolve time: {main_model.Runtime:.2f} seconds"
                result_text.set(result)
                
                # 問題求解成功後提示保存結果按鈕可用
                save_button.config(state=tk.NORMAL)
            else:
                result = "No optimal solution found"
                result_text.set(result)
                save_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            save_button.config(state=tk.DISABLED)

    # 儲存結果為 Excel
    def save_result_to_excel():
        try:
            if main_model is None or log_output is None:
                messagebox.showwarning("Warning", "沒有可保存的模型結果")
                return

            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                save_to_excel(main_model, log_output, file_path)
                messagebox.showinfo("Info", "結果已成功保存")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # 文件路徑變數
    model_file = tk.StringVar()

    # 結果變數
    result_text = tk.StringVar()

    # 建立UI元件
    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack(fill=tk.BOTH, expand=True)

    label = tk.Label(frame, text="模型文件路徑:")
    label.grid(row=0, column=0, sticky=tk.W)

    entry = tk.Entry(frame, textvariable=model_file, width=50)
    entry.grid(row=0, column=1, padx=5, pady=5)

    button_browse = tk.Button(frame, text="瀏覽...", command=select_file)
    button_browse.grid(row=0, column=2, padx=5, pady=5)

    button_solve = tk.Button(frame, text="求解模型", command=solve_model)
    button_solve.grid(row=1, column=1, padx=5, pady=5, sticky=tk.E)

    result_label = tk.Label(frame, textvariable=result_text, justify=tk.LEFT)
    result_label.grid(row=2, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)

    save_button = tk.Button(frame, text="保存結果到 Excel", command=save_result_to_excel, state=tk.DISABLED)
    save_button.grid(row=3, column=1, padx=5, pady=5, sticky=tk.E)

    root.mainloop()

if __name__ == "__main__":
    main()
