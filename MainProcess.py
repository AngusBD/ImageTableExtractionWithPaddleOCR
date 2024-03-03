import cv2
import os
from PaddleOCR import recognize_text
from DataToExcel import data_to_excel

# 檢查目錄
slice_dir = 'SlicedImages'
if not os.path.exists(slice_dir):
    os.makedirs(slice_dir)

# 加載影像
input_dir = 'InputImages'
print("InputImages: ", os.path.join(input_dir,'testImg.jpg'))
image = cv2.imread(os.path.join(input_dir,'testImg.jpg'))

# 轉換灰階
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# 閾值處理
_, thresh = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)

# 找到輪廓
contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

# 儲存圖像座標
images_info = []

# 篩選矩形輪廓
for contour in contours:
    perimeter = cv2.arcLength(contour, True)
    approx = cv2.approxPolyDP(contour, 0.02 * perimeter, True)
    if len(approx) == 4:  # 4個頂點為矩形
        # 獲取邊界與座標
        x, y, w, h = cv2.boundingRect(contour)
        # 計算面積
        area = w * h
        images_info.append((x, y, w, h))

# 排除離群值(忽略面積太小者)
if images_info:
    median_area = sorted([w * h for x, y, w, h in images_info])[len(images_info) // 2]
    outlier_area = median_area / 2
    images_info = [(x, y, w, h) for x, y, w, h in images_info if w * h >= outlier_area]

# 依圖像座標對圖像排序
images_info.sort(key=lambda info: (info[1], info[0]))

# 圖片計數
counter = 0


# 保存非離群值影像
for x, y, w, h in images_info:
    # 裁剪影像
    cropped_image = image[y:y+h, x:x+w]
    # 建立輸出路徑
    counter_str = str(counter).zfill(2)
    output_path = os.path.join(slice_dir, f'cropped_{counter_str}_x{x}_y{y}.jpg')
    # 保存影像
    cv2.imwrite(output_path, cropped_image)
    counter += 1

print(f'{counter} images saved to {slice_dir}')

# 影像座標文檔(開發用)
info_file = os.path.join(slice_dir, 'images_info.txt')
with open(info_file, 'w') as f:
    for info in images_info:
        f.write(','.join(map(str, info)) + '\n')

print(f'Image info saved to {info_file}')


txts = recognize_text(slice_dir)
# txts = ['Description', 'Specification', 'Tolerance', 'L1', 'L2', 'L3', 'Min.Pitch', '8', '26', '-', 'FinishedBottom', '25', '26', '25', 'L1:+/-3', 'FinishedBottom', '8', '26', '一', 'MAX', 'Recessed', '8', None, None, None, 'Min.Trace', '26', 'FinishedTrace', '7', '+/-6', 'FinishedTrace', '26', '+/-6', 'Min.BumpPitch', '30']

# print(images_info)
# print(txts)
data_to_excel(images_info, txts)