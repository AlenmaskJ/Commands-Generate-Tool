import pandas as pd
import os


def generate_commands(excel_path, msisdn_col, smsrouterid_col, output_txt="generated_commands.txt"):
    """
    根据Excel中的数据生成指定格式的命令

    参数：
    excel_path: Excel文件路径（如：data.xlsx）
    msisdn_col: MSISDN对应的列名（如：A列填"MSISDN"，则传"MSISDN"）
    smsrouterid_col: SMSRouterID对应的列名（如：B列填"SMSRouterID"，则传"SMSRouterID"）
    output_txt: 生成命令的保存文件路径，默认是generated_commands.txt
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path, engine="openpyxl")

        # 检查指定的列是否存在
        if msisdn_col not in df.columns:
            raise ValueError(f"Excel中未找到列名：{msisdn_col}")
        if smsrouterid_col not in df.columns:
            raise ValueError(f"Excel中未找到列名：{smsrouterid_col}")

        # 存储生成的命令
        commands = []

        # 遍历每一行数据生成命令
        for index, row in df.iterrows():
            x = row[msisdn_col]
            y = row[smsrouterid_col]

            # 跳过空值行
            if pd.isna(x) or pd.isna(y):
                print(f"第{index + 2}行数据为空，已跳过（Excel表头为第1行）")
                continue

            # 拼接命令
            command = f"Mod Bscex:MSISDN={x},SMSRouterID={y};"
            commands.append(command)

        # 打印生成的命令
        print("===== 生成的命令如下 =====")
        for cmd in commands:
            print(cmd)

        # 将命令保存到txt文件
        with open(output_txt, "w", encoding="utf-8") as f:
            f.write("\n".join(commands))

        print(f"\n所有命令已保存到：{os.path.abspath(output_txt)}")
        return commands

    except FileNotFoundError:
        print(f"错误：未找到Excel文件 {excel_path}，请检查文件路径是否正确")
    except Exception as e:
        print(f"执行出错：{str(e)}")


# ---------------------- 配置区域（请根据你的Excel文件修改以下参数） ----------------------
if __name__ == "__main__":
    # 配置项说明：
    # 1. excel_file：你的Excel文件路径（如果Excel和脚本在同一文件夹，直接填文件名，如"数据.xlsx"）
    # 2. msisdn_column：Excel中MSISDN（x值）所在列的列名（比如你Excel里这一列的表头是"号码"，就填"号码"）
    # 3. smsrouterid_column：Excel中SMSRouterID（y值）所在列的列名（比如表头是"路由ID"，就填"路由ID"）
    excel_file = "data.xlsx"  # 替换为你的Excel文件路径
    msisdn_column = "MSISDN"  # 替换为你Excel中x值列的列名
    smsrouterid_column = "SMSRouterID"  # 替换为你Excel中y值列的列名

    # 调用函数生成命令
    generate_commands(excel_file, msisdn_column, smsrouterid_column)