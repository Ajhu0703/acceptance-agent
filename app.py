import streamlit as st
import docx
from docx.shared import Inches
import pdfplumber
import io
import re

st.set_page_config(page_title="材料驗收單自動生成器", layout="centered")

st.title("📋 材料驗收單自動生成器")

# 1. 基本資訊填寫
col1, col2 = st.columns(2)
with col1:
    class_name = st.text_input("班級名稱", value="115W0149-飲調輕食暨餐飲創業(桃園)-第1期")
with col2:
    date_str = st.text_input("進貨日期", value="115/02/03")

st.markdown("---")

# 2. 檔案上傳區
uploaded_pdf = st.file_uploader("1. 上傳材料明細 PDF (.pdf)", type=["pdf"])
uploaded_doc = st.file_uploader("2. 上傳空白 Word 範本 (.docx)", type=["docx"])
uploaded_images = st.file_uploader("3. 上傳材料照片 (多張一次選取)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

def parse_pdf_items(pdf_file):
    """精準解析 PDF 表格取得品號、品名、規格、數量"""
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

# 3. 產生文件按鈕
if st.button("🚀 產生驗收單文件"):
    if not uploaded_doc or not uploaded_pdf:
        st.error("請確保上傳了 Word 範本與 PDF 明細檔！")
    else:
        # A. 解析 PDF 材料清單
        items_dict = parse_pdf_items(uploaded_pdf)
        
        # B. 智慧處理照片對應（自動依檔名內的數字排序，例如 1 號圖對應品號 1）
        sorted_images = []
        if uploaded_images:
            def extract_img_number(img_file):
                nums = re.findall(r'\d+', img_file.name)
                return int(nums[-1]) if nums else 999
            
            sorted_images = sorted(uploaded_images, key=extract_img_number)

        # C. 載入 Word 範本
        doc = docx.Document(uploaded_doc)

        img_idx = 0
        # D. 遍歷 Word 表格，填入文字與照片
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()

                    # 1. 填寫頂端抬頭
                    if "驗收單號" in cell_text or "班級" in cell_text or "115W0149" in cell_text:
                        cell.text = class_name
                    elif "進貨日" in cell_text:
                        cell.text = f"進貨日：{date_str}"

                    # 2. 處理照片格 (儲存格內有「照片」或是空的純圖片預留格)
                    elif "照片" in cell_text:
                        cell.text = "" # 清空標籤
                        if img_idx < len(sorted_images):
                            p = cell.paragraphs[0]
                            p.add_run().add_picture(sorted_images[img_idx], width=Inches(2.2))
                            img_idx += 1

                    # 3. 處理文字格 (關鍵修正：找尋「品號」加上數字)
                    elif "品號" in cell_text:
                        # 抓出「品號1」、「品號 1」裡面的數字
                        match = re.search(r'品號\s*(\d+)', cell_text)
                        if match:
                            p_num = match.group(1)
                            if p_num in items_dict:
                                info = items_dict[p_num]
                                cell.text = f"品號{p_num} {info['name']} {info['spec']} 數量:{info['qty']}"

        # E. 末端加上會驗結果聲明
        p = doc.add_paragraph()
        p.add_run("會驗結果：經確認，到貨物品均於效期內，且與規格相符。").bold = True

        # F. 輸出下載
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("🎉 驗收單文件生成成功！")
        st.download_button(
            label="📥 下載完成的 Word 檔",
            data=bio.getvalue(),
            file_name=f"驗收紀錄_{class_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
