import pandas as pd


def main():
    # 读取 kb10.xlsx 文件
    df_kb = pd.read_excel('kb10.xlsx')
    # 读取 group10.xlsx 文件
    df_group = pd.read_excel('group10.xlsx')
    # 为两个 DataFrame 添加来源标记
    df_kb['来源'] = 'kb10.xlsx'
    df_group['来源'] = 'group10.xlsx'
    # 按照 ODM、Suppliers、KB Spec、Supllier QTY 列进行合并，indicator=True 会添加一个新列 _merge 表示数据来源
    merged_df = pd.merge(df_kb, df_group, on=['ODM', 'Suppliers', 'KB Spec', 'Supllier QTY'], how='outer', indicator=True)
    # 找出不同的行，即 _merge 列中不为 both 的行
    different_rows = merged_df[merged_df['_merge']!= 'both']
    # 只保留数据，不保留 _merge 列
    different_rows = different_rows.drop(columns=['_merge'])
    # 将结果存储到新的 excel 文件中
    different_rows.to_excel('不同行结果10月.xlsx', index=False)


if __name__ == "__main__":
    main()