import pandas as pd
import random

def load_experts_from_xlsx(file_path):
    """
    从 XLSX 文件加载专家名单。

    参数:
        file_path (str): XLSX 文件的路径。

    返回:
        pandas.DataFrame: 包含专家信息的 DataFrame，如果文件读取失败则返回 None。
    """
    try:
        # 读取 Excel 文件
        df = pd.read_excel(file_path)
        print(f"成功从 '{file_path}' 文件中加载了 {len(df)} 位专家信息。")
        # 打印列名，方便用户了解文件包含哪些信息
        print(f"文件包含的列名: {df.columns.tolist()}")
        return df
    except FileNotFoundError:
        print(f"错误：文件 '{file_path}' 未找到。请检查文件路径是否正确。")
        return None
    except Exception as e:
        print(f"读取文件时发生错误: {e}")
        return None

def extract_random_experts(experts_df, num_to_extract):
    """
    从专家 DataFrame 中随机抽取指定数量的专家。

    参数:
        experts_df (pandas.DataFrame): 包含专家信息的 DataFrame。
        num_to_extract (int): 需要抽取的专家数量。

    返回:
        pandas.DataFrame: 包含被抽取专家信息的 DataFrame。如果无法抽取，则返回 None 或空 DataFrame。
    """
    if experts_df is None or experts_df.empty:
        print("错误：专家名单为空，无法进行抽取。")
        return None

    num_available_experts = len(experts_df)

    if num_to_extract <= 0:
        print("错误：抽取的专家数量必须大于 0。")
        return None
    elif num_to_extract > num_available_experts:
        print(f"警告：要抽取的专家数量 ({num_to_extract}) 大于名单中的专家总数 ({num_available_experts})。")
        print("将返回所有专家。")
        return experts_df.copy() # 返回所有专家的副本
    else:
        # 使用 pandas 的 sample 方法进行随机抽样
        extracted_experts = experts_df.sample(n=num_to_extract, random_state=random.randint(1, 1000)) # random_state 用于可复现性，可选
        return extracted_experts

def display_experts(experts_df, title="抽取的专家名单"):
    """
    在控制台显示专家信息。

    参数:
        experts_df (pandas.DataFrame): 包含专家信息的 DataFrame。
        title (str): 显示时的标题。
    """
    if experts_df is not None and not experts_df.empty:
        print(f"\n--- {title} ---")
        # 使用 to_string() 以便更好地在控制台显示整个 DataFrame
        print(experts_df.to_string(index=False))
    elif experts_df is not None and experts_df.empty:
        print(f"\n--- {title} ---")
        print("没有抽取到任何专家。")


def main():
    """
    主函数，执行专家抽取流程。
    """
    print("--- 欢迎使用专家抽取系统 ---")

    # 获取 XLSX 文件路径
    file_path = input("请输入专家名单 XLSX 文件的完整路径 (例如: D:\\data\\experts.xlsx 或 /path/to/experts.xlsx): ")
    experts_data = load_experts_from_xlsx(file_path)

    if experts_data is not None:
        while True:
            try:
                num_str = input(f"当前共有 {len(experts_data)} 位专家。请输入需要抽取的专家数量: ")
                num_to_extract = int(num_str)
                if num_to_extract > 0:
                    break
                else:
                    print("抽取的专家数量必须是正整数。")
            except ValueError:
                print("输入无效，请输入一个数字。")

        # 进行专家抽取
        extracted_experts_df = extract_random_experts(experts_data, num_to_extract)

        # 显示抽取结果
        if extracted_experts_df is not None:
            display_experts(extracted_experts_df)

            # 可选：保存抽取结果到新的 Excel 文件
            save_results = input("\n是否将抽取的专家名单保存到新的 Excel 文件? (输入 'y' 表示是，其他任意键表示否): ").strip().lower()
            if save_results == 'y':
                output_file_path = input("请输入输出文件的名称 (例如: extracted_experts_list.xlsx): ")
                try:
                    extracted_experts_df.to_excel(output_file_path, index=False)
                    print(f"抽取的专家名单已成功保存到 '{output_file_path}'")
                except Exception as e:
                    print(f"保存文件时发生错误: {e}")
        else:
            print("未能完成专家抽取。")

    print("\n--- 感谢使用专家抽取系统 ---")

if __name__ == "__main__":
    main()
