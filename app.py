import streamlit as st
import docx
from docx.shared import Inches
import pdfplumber
from google import genai
from PIL import Image
import re
import io

st.set_page_config(page_title="材料驗收單自動生成器 (AI 智慧版)", layout="centered")

st.title("📋 材料驗收單自動生成器 (AI 智慧版)")

# 1. 基本資訊與 API Key 填寫
col1, col2 = st.columns(2)
with col1:
    class_name = st.text_input("班級名稱", value="115W0149-飲調輕食暨餐飲創業(桃園)-第1期")
with col2:
    date_str = st.text_input("進貨日期", value="115/02/03")

api_key = st.text_input("輸入 Gemini API Key (用於照片 AI 智慧辨識)", type="password")

st.markdown("---")

# 2. 檔案上傳區
uploaded_pdf = st.file_uploader("1. 上傳材料明細 PDF (.pdf)", type=["pdf"])
uploaded_doc = st.file_uploader("2. 上傳空白 Word 範本 (.docx)", type=["docx"])
uploaded_images = st.file_uploader("3. 上傳材料照片 (任意檔名、隨機順序上傳即可)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

def parse_pdf_items(pdf_file):
    """精準解析 PDF 表格結構"""
    items = {}
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or "品號" in str(row[0]):
                        continue
                    clean_row = [str(c).replace('\n', ' ').strip() if c else "" for c in row]
                    if clean_row[0].isdigit():
                        p_num = clean_row[0]
                        items[p_num] = {
                            "name": clean_row[1] if len(clean_row) > 1 else "",
                            "spec": clean_row[2] if len(clean_row) > 2 else "",
                            "qty": clean_row[4] if len(clean_row) > 4 else "1"
                        }
    return items

def match_image_with_ai(client, img_file, items_dict):
    """利用 Gemini 視覺模型辨識照片，並自動匹配最符合的品號"""
    try:
        image = Image.open(img_file)
        # 建立可供比對的清單描述
        items_summary = "\n".join([f"品號 {k}: {v['name']} ({v['spec']})" for k, v in items_dict.items()])
        
        prompt = f"""
        請分析這張照片中的物品或包裝上的文字標籤。
        對照以下材料清單，判斷這張照片最有可能屬於哪一個品號？

        【材料清單】
        {items_summary}

        請僅輸出對應的品號數字（例如：1、2、3...），不要包含任何其他文字。
        """
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[image, prompt]
        )
        
        matched_num = response.text.strip()
        numbers = re.findall(r'\d+', matched_num)
        return numbers[0] if numbers else None
    except Exception as e:
        st.warning(f"照片 AI 辨識失敗: {e}")
        return None

# 3. 產生文件按鈕
if st.button("🚀 產生驗收單文件"):
    if not uploaded_doc or not uploaded_pdf:
        st.error("請上傳 Word 範本與 PDF 明細檔！")
    elif not api_key:
        st.error("請輸入 Gemini API Key 以啟動照片 AI 自動辨識配對！")
    else:
        with st.spinner("🤖 AI 正在辨識照片內容並自動配對材料品號中..."):
            # A. 解析 PDF 材料清單
            items_dict = parse_pdf_items(uploaded_pdf)
            
            # B. AI 照片辨識與自動配對
            client = genai.Client(api_key=api_key)
            image_map = {}
            if uploaded_images:
                for img in uploaded_images:
                    matched_pnum = match_image_with_ai(client, img, items_dict)
                    if matched_pnum and matched_pnum in items_dict:
                        image_map[matched_pnum] = img

            # C. 載入 Word 範本與填寫
            doc = docx.Document(uploaded_doc)

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_text = cell.text.strip()

                        if "驗收單號" in cell_text or "班級" in cell_text:
                            cell.text = class_name
                        elif "進貨日" in cell_text:
                            cell.text = f"進貨日：{date_str}"

                        elif "品號" in cell_text:
                            match = re.search(r'品號\s*(\d+)', cell_text)
                            if match:
                                p_num = match.group(1)
                                
                                # 照片格填入
                                if "照片" in cell_text:
                                    cell.text = ""
                                    if p_num in image_map:
                                        p = cell.paragraphs[0]
                                        p.add_run().add_picture(image_map[p_num], width=Inches(2.2))
                                
                                # 文字資料格填入
                                else:
                                    if p_num in items_dict:
                                        info = items_dict[p_num]
                                        cell.text = f"品號{p_num} {info['name']} {info['spec']} 數量:{info['qty']}"

            # D. 末端加上會驗結果聲明
            p = doc.add_paragraph()
            p.add_run("會驗結果：經確認，到貨物品均於效期內，且與規格相符。").bold = True

            # E. 輸出下載
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("🎉 AI 智慧辨識完成！照片已自動與品號、品名精準配對並填寫完畢！")
            st.download_button(
                label="📥 下載完成的 Word 檔",
                data=bio.getvalue(),
                file_name=f"驗收紀錄_{class_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
