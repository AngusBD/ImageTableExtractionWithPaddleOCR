import cv2
import os
from paddleocr import PaddleOCR

def recognize_text(image_folder):
    # 初始化 PaddleOCR
    ocr = PaddleOCR()
    txts = []

    # 儲存識別結果的列表
    results = []

    # 遍歷圖片資料夾中的影像
    for file in os.listdir(image_folder):
        if file.endswith('.jpg'):
            # 讀取影像
            image_path = os.path.join(image_folder, file)
            image = cv2.imread(image_path)
            
            # 進行文字識別
            result = ocr.ocr(image)
            
            # 輸出識別結果
            results.append(result)
    for result in results:
        for line in result:
            if line == None:
                txts.append(None)
            else:
                txts.append(line[0][1][0])
    return txts
