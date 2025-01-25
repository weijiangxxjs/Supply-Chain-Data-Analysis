import pandas as pd


# 读取 Excel 文件
excel_file_path = 'D:/python_project/lessTwoDays/KB Weekly BPS summary.xlsx'  # 替换为实际的 Excel 文件路径
df = pd.read_excel(excel_file_path, sheet_name='Weekly BPS')


# 定义一个函数来解析和重新格式化日期字符串
def format_shipment_schedule(date_str):
    if isinstance(date_str, str):
        parts = date_str.split(',')
        formatted_parts = []
        for part in parts:
            if '@' in part:
                num, date_part = part.split('@')
                try:
                    date_obj = pd.to_datetime(date_part)
                    formatted_date = date_obj.strftime('%Y%m%d')
                    formatted_parts.append(f"{num}*{formatted_date}")
                except ValueError:
                    # 如果日期转换失败，保留原格式
                    formatted_parts.append(part)
            else:
                formatted_parts.append(part)
        return ','.join(formatted_parts)
    else:
        return date_str


# 应用函数到 Shipment Schedule 列
df['Shipment Schedule'] = df['Shipment Schedule'].apply(format_shipment_schedule)


# 确保所需列已经存在，如果不存在则创建并初始化为 0
if '小于等于2天的数量' not in df.columns:
    df['小于等于2天的数量'] = 0
if '大于2天的数量' not in df.columns:
    df['大于2天的数量'] = 0
if '小于等于10天的数量' not in df.columns:
    df['小于等于10天的数量'] = 0
if '大于10天数量' not in df.columns:
    df['大于10天数量'] = 0
if '大于2天小于等于10天的数量' not in df.columns:
    df['大于2天小于等于10天的数量'] = 0


# 定义一个函数用于将日期字符串转换为 datetime 类型，方便后续日期计算
def str_to_datetime(date_str):
    try:
        return pd.to_datetime(date_str)
    except ValueError:
        return None


# 定义函数用于处理每行数据，计算对应列的值
def process_row(row):
    bps_qty = row['BPS Qty']
    shipment_schedule = row['Shipment Schedule']
    version_date = str_to_datetime(row['Version'])
    if pd.isnull(version_date):
        return row
    sum_qty = 0
    accu_qty_leq_2days = 0
    accu_qty_leq_10days = 0
    parts = shipment_schedule.split(',')
    # 将包含数量和日期的部分提取出来，组成列表，每个元素是 (数量, 日期) 形式的元组
    schedule_parts = [(int(part.split('*')[0]), str_to_datetime(part.split('*')[1]))
                    for part in parts if '*' in part]
    # 按照日期从小到大排序
    schedule_parts.sort(key=lambda x: x[1])

    for qty, date in schedule_parts:
        sum_qty += qty
        date_diff = (date - version_date).days
        if date_diff <= 2:
            accu_qty_leq_2days += qty
        if date_diff <= 10:
            accu_qty_leq_10days += qty
        if sum_qty >= bps_qty:
            break

    # 当累计的小于等于 2 天的数量大于等于 BPS Qty 时，使用 BPS Qty 更新
    row['小于等于2天的数量'] = bps_qty if accu_qty_leq_2days >= bps_qty else accu_qty_leq_2days
    # 当累计的小于等于 10 天的数量大于等于 BPS Qty 时，使用 BPS Qty 更新
    row['小于等于10天的数量'] = bps_qty if accu_qty_leq_10days >= bps_qty else accu_qty_leq_10days
    # 计算大于 2 天的数量
    row['大于2天的数量'] = max(0, bps_qty - row['小于等于2天的数量'])
    # 计算大于 10 天的数量
    row['大于10天数量'] = max(0, bps_qty - row['小于等于10天的数量'])
    # 计算大于 2 天小于等于 10 天的数量
    row['大于2天小于等于10天的数量'] = max(0, row['大于2天的数量'] - row['大于10天数量'])
    return row


# 对每一行应用处理函数
df = df.apply(process_row, axis=1)


# 将指定列设置为百分比格式，保留两位小数
df[['BPS(%)', 'BPS%(<=2)', 'BPS%(>2)']] = df[['BPS(%)', 'BPS%(<=2)', 'BPS%(>2)']].applymap(lambda x: "{:.2%}".format(x))


# 将指定列设置为整数格式
# df[['小于等于2天的数量', '大于2天的数量', '小于等于10天的数量', '大于10天数量', '大于2天小于等于10天的数量']] = df[['小于等于2天的数量', '大于2天的数量', '小于等于10天的数量', '大于10天数量', '大于2天小于等于10天的数量']].astype(int)


# 保存修改后的数据到原 Excel 文件的原工作表
with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Weekly BPS', index=False)