import streamlit as st
import cv2
import numpy as np
import openpyxl
from PIL import Image
import io

# --- 核心閱卷邏輯 (保留您原始的影像處理算法) ---
def preprocess_image(image_bytes):
    # 將上傳的檔案轉為 OpenCV 格式
    file_bytes = np.asarray(bytearray(image_bytes.read()), dtype=np.uint8)
    img = cv2.imdecode(file_bytes, 1)
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    binary = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                                   cv2.THRESH_BINARY, 51, 2)
    return img, binary

def judge_bubble(bin_img, x, y, w=30, h=24):
    roi = bin_img[y:y+h, x:x+w]
    black_count = np.sum(roi == 0)
    return 1 if (black_count / (w * h)) >= 0.5 else 0

# --- Streamlit 網頁介面 ---
st.set_page_config(page_title="高中科學閱卷系統", layout="wide")

st.title("🧪 高中科學閱卷系統 (網頁版)")
st.write("請上傳答案卡照片與 Excel 範本開始閱卷")

# 側邊欄：設定規則
with st.sidebar:
    st.header("⚙️ 閱卷規則設定")
    q1_range = st.text_input("第一大題題號範圍 (如: 1,10)", "1,10")
    q1_score = st.number_input("第一大題配分", value=5)
    q1_ans = st.text_input("第一大題正確答案", "AAAAABBBBB")
    
    st.divider()
    st.info("註：網頁版不支援儲存到本地 D 槽，處理完請下載結果檔。")

# 第一步：檔案上傳
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 上傳成績登記表 (Excel)")
    excel_file = st.file_uploader("選擇 Excel 範本", type=["xlsx"])

with col2:
    st.subheader("2. 上傳答案卡照片 (JPG)")
    uploaded_images = st.file_uploader("可多選上傳照片", type=["jpg", "jpeg"], accept_multiple_files=True)

# 第二步：開始執行
if st.button("🚀 開始批次閱卷"):
    if excel_file and uploaded_images:
        # 讀取 Excel
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active # 假設處理第一個分頁
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, img_file in enumerate(uploaded_images):
            status_text.text(f"正在處理：{img_file.name}")
            
            # 1. 影像處理
            original_img, binary_img = preprocess_image(img_file)
            
            # 2. 這裡插入您的座標計算與辨識邏輯 (items[i] 等)
            # 範例辨識結果
            detected_score = 85  
            seat_no = idx + 1 # 這裡應從辨識結果中取得
            
            # 3. 寫入 Excel 記憶體中
            sheet.cell(row=seat_no + 1, column=3).value = detected_score
            
            # 更新進度條
            progress_bar.progress((idx + 1) / len(uploaded_images))

        # 第三步：提供下載
        status_text.success("✅ 閱卷完成！")
        
        # 將處理完的 Excel 轉為二進制流供下載
        output = io.BytesIO()
        wb.save(output)
        st.download_button(
            label="💾 下載閱卷成績結果",
            data=output.getvalue(),
            file_name="grading_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("請確保已上傳 Excel 範本與至少一張照片。")
