from openpyxl import Workbook

# images_info = [(9, 13, 393, 67), (404, 13, 437, 31), (843, 13, 133, 67), (404, 46, 149, 34), (555, 46, 163, 34), (720, 46, 122, 34), (9, 81, 393, 45), (404, 81, 149, 45), (555, 81, 163, 45), (720, 81, 121, 45), (9, 128, 393, 45), (404, 128, 149, 45), (555, 128, 163, 45), (720, 128, 121, 45), (843, 128, 133, 45), (9, 175, 393, 44), (404, 175, 149, 44), (555, 175, 163, 44), (720, 175, 121, 44), (843, 175, 133, 44), (9, 221, 393, 43), (404, 221, 149, 43), (555, 221, 163, 43), (720, 221, 121, 43), (843, 221, 133, 43), (9, 266, 393, 40), (404, 266, 437, 40), (9, 307, 393, 40), (404, 307, 437, 40), (843, 307, 133, 40), (9, 349, 393, 38), (404, 349, 437, 38), (843, 349, 133, 38), (9, 388, 393, 40), (404, 388, 437, 40)]
# txts = ['Description', 'Specification', 'Tolerance', 'L1', 'L2', 'L3', 'Min.Pitch', '8', '26', '-', 'FinishedBottom', '25', '26', '25', 'L1:+/-3', 'FinishedBottom', '8', '26', '一', 'MAX', 'Recessed', '8', None, None, None, 'Min.Trace', '26', 'FinishedTrace', '7', '+/-6', 'FinishedTrace', '26', '+/-6', 'Min.BumpPitch', '30']

def data_to_excel(images_info, txts):

    txts_filtered = ['None' if item is None else item for item in txts]
    grouped_txts_y = {}
    grouped_txts_x = {}
    # 將 txts 中的字串依照 images_info 中每個元組的第二個值進行分組

    # 建立y軸字典
    for i, item in enumerate(images_info):
        key = item[1]  # 使用每個元組的第二個值作為鍵
        if key in grouped_txts_y:
            grouped_txts_y[key].append(txts_filtered[i])
        else:
            grouped_txts_y[key] = [txts_filtered[i]]
    sorted_grouped_txts_y = dict(sorted(grouped_txts_y.items()))

    # 建立x軸字典
    for i, item in enumerate(images_info):
        key = item[0]  # 使用每個元組的第二個值作為鍵
        if key in grouped_txts_x:
            grouped_txts_x[key].append(txts_filtered[i])
        else:
            grouped_txts_x[key] = [txts_filtered[i]]
    sorted_grouped_txts_x = dict(sorted(grouped_txts_x.items()))

    # 找出x軸個數(建立座標系)
    x_max_key = max(sorted_grouped_txts_x, key=lambda x: len(sorted_grouped_txts_x[x]))
    x_max_value_count = len(sorted_grouped_txts_x[x_max_key])

    # 找出y軸個數(建立座標系)
    y_max_key = max(sorted_grouped_txts_y, key=lambda y: len(sorted_grouped_txts_y[y]))
    y_max_value_count = len(sorted_grouped_txts_y[y_max_key])


    # print(sorted_grouped_txts_x,x_max_value_count)
    # print(sorted_grouped_txts_y,y_max_value_count)

    # 將每個字的x,y與字合併
    combined_data = [(item[0], item[1], txt) for item, txt in zip(images_info, txts_filtered)]

    # print(combined_data)


    # 建立座標系
    x_values = list(sorted_grouped_txts_x.keys())
    y_values = list(sorted_grouped_txts_y.keys())
    coordinates = [['N/A'] * len(x_values) for _ in range(len(y_values))]

    # 填充座標系
    for x, y, data in combined_data:
        x_index = x_values.index(x)
        y_index = y_values.index(y)
        coordinates[y_index][x_index] = data

    wb = Workbook()
    ws = wb.active
    for row in coordinates:
        ws.append(row)

    # 保存工作簿
    wb.save('SavedText.xlsx')

    for row in coordinates:
        print(row)
