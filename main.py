import gurobipy as gp
from gurobipy import GRB
import io
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

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
    print('')
    print("\n".join(["-"*80, "Gurobi Output Log", "-"*80]))
    for line in log_output:
        print(line)
    print("-"*80 + "\n")

    return model, log_output

def save_to_excel(results: dict, filename: str) -> None:
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for model_name, (model, log_output) in results.items():
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

            # 捕獲的日誌輸出，分段保存
            df_log = pd.DataFrame({'Log': log_output})

            model_name_short = model_name[:10]  # 限制名稱長度
            # 寫入數據到不同的工作表
            df.to_excel(writer, sheet_name=f'{model_name_short}_Solution', index=False)
            df_info.to_excel(writer, sheet_name=f'{model_name_short}_Info', index=False)
            df_log.to_excel(writer, sheet_name=f'{model_name_short}_Log', index=False)

def main():
    # 建立主窗口
    root = tk.Tk()
    root.title("Gurobi Model Runner")
    # 禁用視窗拉伸
    root.resizable(False, False)

    results = {}
    file_entries = []

    def add_file_entry():
        frame = tk.Frame(root, padx=10, pady=5)
        frame.grid(row=len(file_entries) + 2, column=0, sticky=tk.W, columnspan=3)
        
        model_file = tk.StringVar()
        file_entries.append(model_file)

        label = tk.Label(frame, text="模型文件路徑:")
        label.grid(row=0, column=0, sticky=tk.W)

        entry = tk.Entry(frame, textvariable=model_file, width=50)
        entry.grid(row=0, column=1, padx=5, pady=5)

        button_browse = tk.Button(frame, text="瀏覽...", command=lambda: select_file(model_file))
        button_browse.grid(row=0, column=2, padx=5, pady=5)

    # 文件選擇函數
    def select_file(model_file):
        file_path = filedialog.askopenfilename(filetypes=[("Model Files", "*.lp *.mps *.json")])
        if file_path:
            model_file.set(file_path)

    # 執行模型求解
    def solve_models():
        try:
            model_file_paths = [entry.get() for entry in file_entries if entry.get()]
            if not model_file_paths:
                messagebox.showwarning("Warning", "請選擇至少一個模型文件")
                return

            results.clear()

            for file_path in model_file_paths:
                model_name = file_path.split("/")[-1]
                model, log_output = execute_model_from_file(file_path)
                results[model_name] = (model, log_output)
            
            if all(model.status == GRB.OPTIMAL for model, _ in results.values()):
                result = "\n".join([f"{name}: Optimal solution: {model.ObjVal}, Solve time: {model.Runtime:.2f} seconds"
                                    for name, (model, _) in results.items()])
                result_text.set(result)

                # 問題求解成功後提示保存結果按鈕可用
                save_button.config(state=tk.NORMAL)
            else:
                result = "Not all models found an optimal solution"
                result_text.set(result)
                save_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            save_button.config(state=tk.DISABLED)
        
        # 清空路徑框
        for entry in file_entries:
            entry.set("")

    # 儲存結果為 Excel
    def save_result_to_excel():
        try:
            if not results:
                messagebox.showwarning("Warning", "沒有可保存的模型結果")
                return

            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                save_to_excel(results, file_path)
                messagebox.showinfo("Info", "結果已成功保存")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # 結果變數
    result_text = tk.StringVar()

    # 建立UI元件
    

    result_label = tk.Label(root, textvariable=result_text, justify=tk.LEFT)
    result_label.grid(row=0, column=0, columnspan=3)

    button_solve = tk.Button(root, text="求解模型", command=solve_models)
    button_solve.grid(row=1, column=0)

    save_button = tk.Button(root, text="保存結果到 Excel", command=save_result_to_excel, state=tk.DISABLED)
    save_button.grid(row=1, column=1)

    button_add_file = tk.Button(root, text="添加模型文件", command=add_file_entry)
    button_add_file.grid(row=1, column=2)

    add_file_entry()  # 添加第一個文件輸入框

    root.mainloop()

if __name__ == "__main__":
    main()
