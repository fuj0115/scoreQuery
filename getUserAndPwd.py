import pandas as pd

# 配置路径（根据实际情况修改）
input_path = "D:\getScore\蔡老师-学员成绩表3月18日.xlsx"  # 替换为你的Excel路径
output_path = "D:\getScore\输出结果.xlsx"  # 新文件名

# 创建空DataFrame容器
combined_data = pd.DataFrame()

# 读取原始文件
with pd.ExcelFile(input_path) as xls:
    # 遍历所有sheet
    for sheet_name in xls.sheet_names:
        # 读取当前sheet数据（指定准考证号和密码列为字符串类型）
        df = pd.read_excel(
            xls, sheet_name=sheet_name, dtype={"准考证号": str, "密码": str}
        )

        # 验证列是否存在
        if all(col in df.columns for col in ["准考证号", "密码"]):
            # 提取目标列并添加来源标记
            temp_df = df[["准考证号", "密码"]].copy()
            temp_df["来源Sheet"] = sheet_name  # 添加来源信息列

            # ---------------------- 新增过滤逻辑 ----------------------
            # 生成处理后的临时列用于条件判断（不修改原始数据）
            cond_kh = temp_df["准考证号"].fillna("").str.strip() == ""
            cond_pwd = temp_df["密码"].fillna("").str.strip() == ""
            cond_kh_forecast = temp_df["准考证号"].fillna("").str.contains("预报")
            cond_pwd_forecast = temp_df["密码"].fillna("").str.contains("预报")

            # 组合条件：任意一个条件满足则过滤掉该行
            mask = cond_kh | cond_pwd | cond_kh_forecast | cond_pwd_forecast
            temp_df = temp_df[~mask]
            # --------------------------------------------------------

            # 追加到合并数据
            combined_data = pd.concat([combined_data, temp_df], ignore_index=True)
        else:
            print(f"跳过 [{sheet_name}] - 缺少准考证号或密码列")

# 写入新文件（如果存在有效数据）
if not combined_data.empty:
    combined_data.to_excel(
        output_path,
        sheet_name="汇总数据",
        index=False,
        columns=["准考证号", "密码", "来源Sheet"],  # 调整列顺序
    )
    print(f"合并完成！共处理 {len(combined_data)} 条数据（过滤后）")
else:
    print("未找到有效数据，请检查原始文件")

print("输出文件路径:", output_path)
