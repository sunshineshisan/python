import pandas as pd

def calculate_in_third_sheet(file_path):
    try:
        # 读取 Excel 文件
        excel_file = pd.ExcelFile(file_path)
        # 获取第三个工作表（索引从 0 开始）
        df = excel_file.parse(2)

        # 检查列数是否足够
        if df.shape[1] < 5:
            print("错误：第三个工作表中的列数不足 5 列，无法进行计算。")
            return None

        # 将空值替换为 0
        df = df.fillna(0)

        # 计算 (第二列 + 第三列) * 第五列
        result = (df.iloc[:, 1] + df.iloc[:, 2]) * (1-(df.iloc[:, 4]/100))/df.iloc[:, 5]*7.2
        return result
    except FileNotFoundError:
        print(f"错误：未找到文件 {file_path}。")
    except IndexError:
        print("错误：文件中不存在第三个工作表。")
    except Exception as e:
        print(f"发生未知错误：{e}")


if __name__ == "__main__":
    file_path = '1.xls'  # 请替换为你的 Excel 文件路径
    result = calculate_in_third_sheet(file_path)
    if result is not None:
        print(result)
    