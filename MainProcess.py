import cv2
import os
import shutil
import numpy as np
from PaddleOCR import recognize_text
from DataToExcel import data_to_excel

# 檢查目錄
slice_dir = 'SlicedImages'
# 如果資料夾存在，則遞迴地刪除它及其底下的所有文件
if os.path.exists(slice_dir):
    shutil.rmtree(slice_dir)

# 創建資料夾
os.makedirs(slice_dir)

# 加載影像
input_dir = 'InputImages'
image = cv2.imread(os.path.join(input_dir,'img1.jpg'))

# 將圖像放大1.5倍
# image = cv2.resize(image, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)

# 進行銳化
# sharpened_image = cv2.GaussianBlur(image, (0, 0), 3)
# sharpened_image = cv2.addWeighted(image, 1.5, sharpened_image, -0.5, 0)

# # 顯示放大且銳化後的圖像
# cv2.imshow('Sharpened Image', sharpened_image)

# 轉換灰階
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# 閾值處理
_, thresh = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)

inverted_image = cv2.bitwise_not(thresh)

iter = 2
hor = np.array([[1,1,1,1,1,1]])
vertical_lines_eroded_image = cv2.erode(inverted_image, hor, iterations=iter)
vertical_lines_eroded_image = cv2.dilate(vertical_lines_eroded_image, hor, iterations=iter)

ver = np.array([[1],
               [1],
               [1],
               [1],
               [1],
               [1],
               [1]])
horizontal_lines_eroded_image = cv2.erode(inverted_image, ver, iterations=iter)
horizontal_lines_eroded_image = cv2.dilate(horizontal_lines_eroded_image, ver, iterations=iter)

combined_image = cv2.add(vertical_lines_eroded_image, horizontal_lines_eroded_image)

# 找到輪廓
contours, _ = cv2.findContours(combined_image, cv2.RETR_TREE, cv2.CHAIN_APPROX_NONE)
image_copy = image.copy()
cv2.drawContours(image_copy, contours, -1, (0, 255, 0), 3)


# 儲存圖像座標
images_info = []

# 篩選矩形輪廓
for contour in contours:
    perimeter = cv2.arcLength(contour, True)
    approx = cv2.approxPolyDP(contour, 0.02 * perimeter, True)

    # 頂點補償
    if len(approx) == 2:
        # 循環遍歷每個頂點
        x1 = approx[0][0][0]
        y1 = approx[0][0][1]
        x2 = approx[1][0][0]
        y2 = approx[1][0][1]
        if x1 != x2 and y1 != y2:
            # 新增 [x1 y2] 和 [x2 y1]
            new_points = np.array([[x1, y2], [x2, y1]])

            # 將 new_points 轉換成三維陣列
            new_points = np.expand_dims(new_points, axis=1)

            # 垂直堆疊新的點到 approx 中
            approx = np.vstack((approx, new_points))

    if len(approx) == 4:  # 4個頂點為矩形
        for point in approx:
            x, y = point[0]  # 獲取頂點座標
            cv2.circle(image_copy, (x, y), 5, (0, 0, 255), -1)  # 繪製頂點
        
        # 獲取邊界與座標
        x, y, w, h = cv2.boundingRect(contour)
        
        # 計算面積
        area = w * h
        images_info.append((x, y, w, h))

cv2.imshow('drawContours',image_copy)

# 排除離群值(忽略面積太小者)
if images_info:
    median_area = sorted([w * h for x, y, w, h in images_info])[len(images_info) // 2]
    
    min_outlier_area = median_area / 5
    max_outlier_area = median_area * 5
    # print(median_area, min_outlier_area, max_outlier_area)
    
    images_info = [(x, y, w, h) for x, y, w, h in images_info if w * h >= min_outlier_area and w * h <= max_outlier_area]

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
    # output_path = os.path.join(slice_dir, f'cropped_{counter_str}.jpg')
    
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

data_to_excel(images_info, txts)
cv2.waitKey(0)