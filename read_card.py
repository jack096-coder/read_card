import streamlit as st
import cv2
import numpy as np
import openpyxl
from PIL import Image
import io

# --- 1. 核心對照表設定 (由原程式碼平移) ---
CHAR_MAP = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', '*', '$', '%', '#' , '=']
JACK_MAP = [[0]*5 for _ in range(31)]
# 初始化劃記組合 (簡化表示，建議依照原程式補完)
JACK_MAP[0] = [1, 0, 0, 0, 0]; JACK_MAP[1] = [0, 1, 0, 0, 0]; JACK_MAP[2] = [0, 0, 1, 0, 0]
JACK_MAP[3] = [0, 0, 0, 1, 0]; JACK_MAP[4] = [0, 0, 0, 0, 1]; JACK_MAP[30] = [0, 0, 0, 0, 0]

# --- 2. 影像處理與座標計算函數 ---
def process_answer_sheet(uploaded_file, bbb_config):
    # 將上傳檔案轉為 OpenCV 格式
    file_bytes = np.asarray(bytearray(uploaded_file.read()), dtype=np.uint8)
    image = cv2.imdecode(file_bytes, 1)
    if image is None: return None, "格式錯誤"

    # 二值化處理
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    ret, thresh = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY_INV)
    
    # 尋找錨點 (方形定位點)
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    listall = []
    for contour in contours:
        approx = cv2.approxPolyDP(contour, cv2.arcLength(contour, True) * 0.04, True)
        if len(approx) == 4:
            x, y, w, h = cv2.boundingRect(approx)
            if abs(w / h - 1.0) < 0.2: # 檢查長寬比
                listall.append([x, y, w, h])
    
    listall = np.array(listall)
    if len(listall) < 36: return image, "偵測錨點不足 (需 25+11)"

    # 找出出現頻率最高的 X 與 Y 以標定邊界
    most_frequent_x = np.argmax(np.bincount(listall[:, 0]))
    most_frequent_y = np.argmax(np.bincount(listall[:, 1]))
    
    error_threshold = 10
    listy = sorted([d for d in listall if abs(d[0] - most_frequent_x) <= error_threshold], key=lambda y: y[1])
    listx = sorted([d for d in listall if abs(d[1] - most_frequent_y) <= error_threshold], key=lambda x: x[0])

    if len(listy) != 25 or len(listx) != 11:
        return image, f"錨點數量異常 (Y:{len(listy)}, X:{len(listx)})"

    # --- 重要：items 座標計算邏輯 (平移自原程式) ---
    items = [[None] * 11 for _ in range(46)]
    d, w, h = 7, 8, 6 # 偏移量微調參數
    binary_img = cv2.adaptiveThreshold(cv2.GaussianBlur(gray, (3,3), 0), 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 51, 2)

    # 年級、班級、座號座標
    items[0] = [listy[0][1]+d-h, listx[1][0]+d-w, listx[2][0]+d-w, listx[3][0]+d-w]
    for i in range(1, 5):
        items[40+i] = [listy[i][1]+d-h] + [listx[j][0]+d-w for j in range(1, 11)]

    # 題目座標 (1-20 題在左, 21-40 題在右)
    for i in range(1, 21):
        items[i] = [listy[i+4][1]+d-h, listx[0][0]+d-w, listx[1][0]+d-w, listx[2][0]+d-w, listx[3][0]+7-w, listx[4][0]+7-w]
    for i in range(21, 41):
        items[i] = [listy[i+4-20][1]+d-h, listx[6][0]+d-w, listx[7][0]+d-w, listx[8][0]+d-w, listx[9][0]+7-w, listx[10][0]+7-w]

    # --- 判斷劃記與給分 ---
    detected_answers = []
    for i in range(1, 41):
        row_ans = []
        for j in range(1, 6):
            # judge_draw 邏輯
            roi = binary_img[items[i][0]:items[i][0]+24, items[i][1]:items[i][1]+30]
            point = np.sum(roi == 0)
            is_drawn = 1 if (point / 720) >= 0.50 else 0
            row_ans.append(is_drawn)
            if is_drawn: cv2.rectangle(image, (items[i][j], items[i][0]), (items[i][j]+29, items[i][0]+23), (0,0,255), 2)
        detected_answers.append(row_ans)

    # 辨識班號座號 (簡化邏輯)
    grade = 1 # 預設值，實際應由 items[0] 判定
    class_id = 101 # 實際應由 items[41-42] 判定
    seat_id = 1 # 實際應由 items[43-44] 判定

    return image, {"grade": grade, "class": class_id, "seat": seat_id, "ans": detected_answers}

# --- 3. Streamlit 介面佈局 ---
st.set_page_config(page_title="科學閱卷系統 Web", layout="wide")
st.title("📑 自動閱卷與座標辨識系統")

with st.sidebar:
    st.header("📋 閱卷規則 (bbb 設定)")
    q1_start = st.number_input("第一大題起始題號", 1)
    q1_end = st.number_input("第一大題結束題號", 20)
    q1_score = st.number_input("第一大題配分", 5)
    q1_ans_str = st.text_input("第一大題標準答案", "ABCDE...")
    # 將設定包裝成 bbb 列表結構
    bbb_config = [q1_start, q1_end, q1_score, q1_ans_str, None, None, None, None, None, None, None, None]

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. 上傳資料")
    excel_tpl = st.file_uploader("成績登記表範本", type=["xlsx"])
    img_files = st.file_uploader("答案卡照片 (可多選)", type=["jpg", "png"], accept_multiple_files=True)

if st.button("▶️ 開始批次閱卷"):
    if not img_files or not excel_tpl:
        st.warning("請確保已上傳 Excel 範本與照片檔。")
    else:
        wb = openpyxl.load_workbook(excel_tpl)
        results_summary = []
        
        for up_file in img_files:
            processed_img, data = process_answer_sheet(up_file, bbb_config)
            
            if isinstance(data, dict):
                st.write(f"✅ {up_file.name}: 辨識成功 ({data['grade']}年{data['class']}班 {data['seat']}號)")
                # 這裡寫入 Excel 邏輯 (sheet.cell...)
                # ...
                results_summary.append(data)
                # 顯示辨識後的圈選圖
                st.image(processed_img, caption=f"{up_file.name} 辨識結果", use_container_width=True)
            else:
                st.error(f"❌ {up_file.name}: {data}")

        # 下載結果
        out_io = io.BytesIO()
        wb.save(out_io)
        st.download_button("📥 下載閱卷完成清單", out_io.getvalue(), "graded_results.xlsx")
