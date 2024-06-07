import gurobipy as gp
from gurobipy import GRB
import time
import tkinter as tk
from tkinter import filedialog, messagebox

def execute_model_from_file(model_file: str) -> tuple[gp.Model, float]:
    # 從文件讀取模型
    model = gp.read(model_file)

    # 開始計時
    start_time = time.time()

    # 求解模型
    model.optimize()

    # 計算求解時間
    solve_time = time.time() - start_time

    return model, solve_time

def main():
    # 建立主窗口
    root = tk.Tk()
    root.title("Gurobi Model Runner")

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
            
            main_model, solve_time = execute_model_from_file(model_file_path)
            
            if main_model.status == GRB.OPTIMAL:
                result = f"Optimal solution: {main_model.ObjVal}\nSolve time: {solve_time:.2f} seconds"
            else:
                result = "No optimal solution found"
            
            result_text.set(result)
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

    root.mainloop()

if __name__ == "__main__":
    main()
