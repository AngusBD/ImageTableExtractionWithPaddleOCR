import cv2
import os
from paddleocr import PaddleOCR

def recognize_text(image_folder):
    # 初始化 PaddleOCR
    ocr = PaddleOCR()
    txts = []

    # 存储识别结果的列表
    results = []

    # 遍历图片文件夹中的图像
    for file in os.listdir(image_folder):
        if file.endswith('.jpg'):
            # 读取图像
            image_path = os.path.join(image_folder, file)
            image = cv2.imread(image_path)
            
            # 进行文字识别
            result = ocr.ocr(image)
            
            # 输出识别结果
            results.append(result)
    for result in results:
        for line in result:
            if line == None:
                txts.append(None)
            else:
                txts.append(line[0][1][0])
    return txts

# # 调用函数进行文字识别
# image_folder = 'sliceImg'
# text_results = recognize_text(image_folder)
# for result in text_results:
#     for line in result:
#         if line == None:
#             print(None)
#         else:
#             print(line[0][1][1])