import streamlit as st
import docx
from docx.shared import Inches
import pdfplumber
import io
import re

# 引用 Google AI 套件
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="材料驗收單自動生成器", layout="centered")
st.title("📋 材料驗收單自動生成器 (Gemini 視覺版)")

# --- 新增 API 金鑰輸入框 ---
st.sidebar.title("✨ AI 設定")
google_api_key = st.sidebar.text_input("請貼上您的 Google AI API Key", type="password")

if not google_api_key:
    st.sidebar.warning("請在左側側邊欄輸入您的 API Key 以啟用 AI 辨識功能。")

col1, col2 = st.columns(2)
with col1:
    class_name = st.text_input("班級名稱", value="115W0149-飲調輕食暨餐飲創業(桃園)-第1期")
with col2:
    date_str = st.text_input("進貨日期", value="115/02/03")

st.markdown("---")

uploaded_pdf = st.file_uploader("1. 上傳材料明細 PDF (.pdf)", type=["pdf"])
uploaded_doc = st.file_uploader("2. 上傳空白 Word 範本 (.docx)", type=["docx"])
uploaded_images = st.file_uploader("3. 上傳材料照片 (多張一次選取)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

def parse_pdf_items(pdf_file):
    """從PDF解析品項資料"""
    items = {}
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
                        p_qty = "1"
                        if len(clean_row) >= 6:
                            p_qty = clean_row[4]
                        elif len(clean_row) == 5:
                            p_qty = clean_row[3]
                        items[p_num] = { "name": p_name, "spec": p_spec, "qty": p_qty }
    return items

if st.button("🚀 產生驗收單文件"):
    if not all([uploaded_doc, uploaded_pdf, google_api_key]):
        st.error("請確保上傳了 Word 範本、PDF 明細，並已在左側輸入 API Key！")
    else:
        # A. 解析 PDF
        items_dict = parse_pdf_items(uploaded_pdf)
        
        # 將材料清單格式化為給 AI 看的文字
        materials_text = "這是一份材料清單，包含品號、品名和規格：\n"
        for p_num, info in items_dict.items():
            materials_text += f"- 品號: {p_num}, 品名: {info['name']}, 規格: {info['spec']}\n"

        # --- B. 核心修改：使用 Gemini 視覺辨識來配對照片 ---
        genai.configure(api_key=google_api_key)
        model = genai.GenerativeModel('gemini-1.5-flash') # 使用最新的 Flash 模型，速度快且成本低

        image_map = {}
        unmatched_images = []

        if uploaded_images:
            st.info("🤖 AI 正在辨識照片內容，請稍候...")
            progress_bar = st.progress(0)

            for i, img_file in enumerate(uploaded_images):
                try:
                    image = Image.open(img_file)
                    
                    # 準備給 AI 的指令 (Prompt)
                    prompt_parts = [
                        "任務：請根據下方提供的材料清單，判斷這張圖片中的物品最符合清單中的哪一個項目。\n",
                        "材料清單：\n",
                        materials_text,
                        "\n圖片：\n",
                        image,
                        "\n---",
                        "請只回覆最相符的「品號」數字，不要包含任何其他文字或解釋。"
                    ]

                    # 呼叫 Gemini API
                    response = model.generate_content(prompt_parts)
                    
                    # 從 AI 回應中提取品號數字
                    p_num_match = re.search(r'\d+', response.text)
                    if p_num_match:
                        p_num = p_num_match.group(0)
                        if p_num in items_dict:
                             image_map[p_num] = img_file
                        else:
                             unmatched_images.append(f"{img_file.name} (AI回傳品號 {p_num} 不在清單中)")
                    else:
                        unmatched_images.append(f"{img_file.name} (AI無法辨識)")

                except Exception as e:
                    unmatched_images.append(f"{img_file.name} (處理失敗: {e})")
                
                progress_bar.progress((i + 1) / len(uploaded_images))

            if unmatched_images:
                st.warning(f"注意：有 {len(unmatched_images)} 張照片無法成功配對：\n- " + "\n- ".join(unmatched_images))

        # C. 載入 Word 範本
        doc = docx.Document(uploaded_doc)
        
        # D. 遍歷 Word 表格與替換
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
                                    cell.text = f"品號{p_num} {info['name']} {info['spec']} 數量:{info['qty']}"
        # E. 加會驗結果
        p = doc.add_paragraph()
        p.add_run("會驗結果：經確認，到貨物品均於效期內，且與規格相符。").bold = True

        # F. 輸出下載
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("🎉 驗收單文件生成成功！已透過 Gemini AI 視覺辨識自動對應照片！")
        st.download_button(
            label="📥 下載完成的 Word 檔",
            data=bio.getvalue(),
            file_name=f"驗收紀錄_{class_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
