import gurobipy as gp
from gurobipy import GRB
import time
import subprocess

def check_cuda_support() -> bool:
    try:
        # 使用 nvidia-smi 檢查是否存在 NVIDIA GPU
        subprocess.run(["nvidia-smi"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True
    except subprocess.CalledProcessError:
        return False

def execute_model(use_cuda: bool, model_name: str, model_definition: callable) -> tuple[gp.Model, float]:
    # 創建模型
    model = gp.Model(model_name)

    # 調用外部定義的模型添加變數和約束條件
    model_definition(model)

    # 設置參數
    model.setParam('UseCuda', 1 if use_cuda else 0)

    # 開始計時
    start_time = time.time()

    # 求解模型
    model.optimize()

    # 計算求解時間
    solve_time = time.time() - start_time

    return model, solve_time

def define_test_model(model: gp.Model) -> None:
    # 添加測試模型的變數和約束條件
    x = model.addVar(vtype=GRB.CONTINUOUS, name="x")
    y = model.addVar(vtype=GRB.CONTINUOUS, name="y")
    model.setObjective(3*x + 2*y, GRB.MAXIMIZE)
    model.addConstr(x + y <= 4, "c0")
    model.addConstr(x - y >= 1, "c1")

def define_main_model(model: gp.Model) -> None:
    # 在此處定義你的正式模型
    # 這裡僅作為示例，需要根據實際情況添加
    x = model.addVar(vtype=GRB.CONTINUOUS, name="x")
    y = model.addVar(vtype=GRB.CONTINUOUS, name="y")
    model.setObjective(3*x + 2*y, GRB.MAXIMIZE)
    model.addConstr(x + y <= 4, "c0")
    model.addConstr(x - y >= 1, "c1")
    model.addConstr(x + y >= float(input("Enter a value for x + y: ")), "c2")
    model.addConstr(x - y <= float(input("Enter a value for x - y: ")), "c3")

def main() -> None:
    # 檢查是否支援 CUDA
    cuda_supported = check_cuda_support()
    if cuda_supported:
        device = input("此設備支援 CUDA 加速。請選擇執行設備(cpu:0, gpu:1, 自動依性能選擇:3)，按 Enter 鍵繼續：")
    else:
        device = 0
        print("此設備不支援 CUDA 加速。Running on CPU.")
    
    use_cuda = False
    if device == 0:
        use_cuda = False
    elif device == 1:
        use_cuda = True
    elif device == 3:
        # 自動選擇設備
        # 用測試模型測試並測量時間
        _, time_cpu = execute_model(use_cuda=False, model_name="test_model", model_definition=define_test_model)
        _, time_cuda = execute_model(use_cuda=True, model_name="test_model", model_definition=define_test_model)
        # 比較性能，選擇較快的設定
        if time_cuda < time_cpu:
            print("Using CUDA is faster. Running on GPU.")
            use_cuda = True
        else:
            print("Using CPU is faster. Running on CPU.")
            use_cuda = False

    # 使用選定的設備來求解正式模型
    main_model, solve_time = execute_model(use_cuda=use_cuda, model_name="main_model", model_definition=define_main_model)

    # 輸出結果
    if main_model.status == GRB.OPTIMAL:
        x, y = main_model.getVars()
        print(f"Optimal solution: x = {x.X}, y = {y.X}")
        print(f"Objective value: {main_model.ObjVal}")
        print(f"Solve time: {solve_time}")
    else:
        print("No optimal solution found")

main()