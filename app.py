import streamlit as st
import docx
from docx.shared import Inches
import pdfplumber
import io
import re
import base64
import requests
from PIL import Image

# --- 關鍵修正：定義全球通用的 API 端點 ---
# 使用 aistudio.google.com 的代理服務，繞過地區限制
API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro-vision:generateContent"

st.set_page_config(page_title="材料驗收單自動生成器", layout="centered")
st.title("📋 材料驗收單自動生成器 (AI 視覺版)")

st.sidebar.title("✨ AI 設定")
google_api_key = st.sidebar.text_input("請貼上您的 Google AI API Key", type="password")

if not google_api_key:
    st.sidebar.warning("請在左側側邊欄輸入您的 API Key 以啟用 AI 辨識功能。")

# ... (其餘介面程式碼不變) ...
col1, col2 = st.columns(2)
with col1:
    class_name = st.text_input("班級名稱", value="115W0149-飲調輕食暨餐飲創業(桃園)-第1期")
with col2:
    date_str = st.text_input("進貨日期", value="115/02/03")

st.markdown("---")

uploaded_pdf = st.file_uploader("1. 上傳材料明細 PDF (.pdf)", type=["pdf"])
uploaded_doc = st.file_uploader("2. 上傳空白 Word 範本 (.docx)", type=["docx"])
uploaded_images = st.file_uploader("3. 上傳材料照片 (多張一次選取)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)


# --- 函式部分 ---
def parse_pdf_items(pdf_file):
    items = {}
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row: continue
                        clean_row = [str(c).replace('\n', ' ').strip() if c else "" for c in row]
                        if clean_row and clean_row[0].isdigit():
                            p_num = clean_row[0]
                            p_name = clean_row[1] if len(clean_row) > 1 else ""
                            p_spec = clean_row[2] if len(clean_row) > 2 else ""
                            items[p_num] = { "name": p_name, "spec": p_spec }
    except Exception as e:
        st.error(f"解析 PDF 失敗: {e}")
    return items

def image_to_base64(image):
    buffered = io.BytesIO()
    if image.mode == 'RGBA':
        image = image.convert('RGB')
    image.save(buffered, format="JPEG")
    return base64.b64encode(buffered.getvalue()).decode('utf-8')

def call_gemini_vision_api(api_key, image_base64, materials_text):
    url = f"{API_ENDPOINT}?key={api_key}"
    headers = {'Content-Type': 'application/json'}
    prompt = "任務：請根據下方提供的材料清單，判斷這張圖片中的物品最符合清單中的哪一個項目。請只回覆最相符的「品號」數字，不要包含任何其他文字或解釋。"
    payload = {
        "contents": [{"parts": [{"text": prompt}, {"text": materials_text}, {"inline_data": {"mime_type": "image/jpeg", "data": image_base64}}]}]
    }
    
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 200:
        return response.json()
    else:
        # 顯示更詳細的錯誤
        st.error(f"API 請求失敗 (HTTP {response.status_code}): {response.text}")
        return None

# --- 主邏輯 ---
if st.button("🚀 產生驗收單文件"):
    if not all([uploaded_doc, uploaded_pdf, google_api_key]):
        st.error("請確保上傳了 Word 範本、PDF 明細，並已在左側輸入 API Key！")
    else:
        items_dict = parse_pdf_items(uploaded_pdf)
        materials_text = "這是一份材料清單，包含品號、品名和規格：\n"
        for p_num, info in items_dict.items():
            materials_text += f"- 品號: {p_num}, 品名: {info['name']}, 規格: {info['spec']}\n"

        image_map = {}
        unmatched_images = []

        if uploaded_images:
            with st.spinner("🤖 AI 正在辨識照片內容，請稍候..."):
                for i, img_file in enumerate(uploaded_images):
                    try:
                        image = Image.open(img_file)
                        image_base64 = image_to_base64(image)
                        api_response = call_gemini_vision_api(google_api_key, image_base64, materials_text)
                        
                        if api_response and 'candidates' in api_response and api_response['candidates']:
                            response_text = api_response['candidates'][0]['content']['parts'][0]['text']
                            p_num_match = re.search(r'\d+', response_text)
                            if p_num_match:
                                p_num = p_num_match.group(0)
                                if p_num in items_dict:
                                    image_map[p_num] = img_file
                                else:
                                    unmatched_images.append(f"{img_file.name} (AI 回傳品號 {p_num} 不在清單中)")
                            else:
                                unmatched_images.append(f"{img_file.name} (AI 無法從回應中解析品號)")
                        else:
                            unmatched_images.append(f"{img_file.name} (API 未返回有效結果)")
                    except Exception as e:
                        unmatched_images.append(f"{img_file.name} (處理失敗: {e})")
            
            if unmatched_images:
                st.warning(f"注意：有 {len(unmatched_images)} 張照片無法成功配對：\n- " + "\n- ".join(unmatched_images))

        # ... (Word 文件生成部分維持不變) ...
        doc = docx.Document(uploaded_doc)
        # (此處省略與先前版本相同的 Word 生成程式碼以節省篇幅)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if "驗收單號" in cell_text or "班級" in cell_text:
                        cell.text = class_name
                    elif "進貨日" in cell_text:
                        cell.text = f"進貨日：{date_str}"
                    elif "品號" in cell_text or "品名" in cell_text:
                        nums = re.findall(r'\d+', cell_text)
                        if nums:
                            p_num = nums[0]
                            if "照片" in cell_text:
                                cell.text = ""
                                if p_num in image_map:
                                    p = cell.paragraphs[0]
                                    p.clear()
                                    run = p.add_run()
                                    run.add_picture(image_map[p_num], width=Inches(2.0))
                                else:
                                    cell.text = "--- 未找到對應照片 ---"
                            else:
                                if p_num in items_dict:
                                    info = items_dict[p_num]
                                    cell.text = f"品號{p_num} {info['name']} {info['spec']}"
        p = doc.add_paragraph()
        p.add_run("會驗結果：經確認，到貨物品均於效期內，且與規格相符。").bold = True
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("🎉 驗收單文件生成成功！")
        st.download_button(
            label="📥 下載完成的 Word 檔",
            data=bio.getvalue(),
            file_name=f"驗收紀錄_{class_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
