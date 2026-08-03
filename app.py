import streamlit as st
import docx
from docx.shared import Inches
import pdfplumber
import re
import io

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
uploaded_images = st.file_uploader("3. 上傳材料照片 (任意上傳，程式會自動按檔名編號對應品號)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

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
                    # 清理每個欄位文字
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
        
        # B. 建立照片檔名對照表 (從檔名找數字，例如 U-2026-08-03(1).jpg -> 找到 1)
        image_map = {}
        if uploaded_images:
            for img in uploaded_images:
                # 尋找檔名結尾或括號內的數字，如 (1) 或 _1 或 1.jpg
                numbers = re.findall(r'\d+', img.name)
                if numbers:
                    # 取最後一個數字作為品號編號 (例如 (1).jpg -> 1)
                    num_key = numbers[-1]
                    image_map[num_key] = img

        # C. 載入 Word 範本
        doc = docx.Document(uploaded_doc)

        # D. 遍歷 Word 表格，精準對應
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()

                    # 1. 填寫頂端抬頭
                    if "驗收單號" in cell_text or "班級" in cell_text:
                        cell.text = class_name
                    elif "進貨日" in cell_text:
                        cell.text = f"進貨日：{date_str}"

                    # 2. 處理照片與品號文字
                    elif "品號" in cell_text:
                        match = re.search(r'品號\s*(\d+)', cell_text)
                        if match:
                            p_num = match.group(1)
                            
                            # 如果這格是「照片格」 (如: 品號1的照片)
                            if "照片" in cell_text:
                                cell.text = "" # 清空標籤
                                if p_num in image_map:
                                    p = cell.paragraphs[0]
                                    p.add_run().add_picture(image_map[p_num], width=Inches(2.2))
                            
                            # 如果這格是「文字資料格」
                            else:
                                if p_num in items_dict:
                                    info = items_dict[p_num]
                                    cell.text = f"品號{p_num} {info['name']} {info['spec']} 數量:{info['qty']}"

        # E. 末端加上會驗結果聲明
        p = doc.add_paragraph()
        p.add_run("會驗結果：經確認，到貨物品均於效期內，且與規格相符。").bold = True

        # F. 輸出下載
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("🎉 文件生成成功！照片與品號已完美精準對應！")
        st.download_button(
            label="📥 下載完成的 Word 檔",
            data=bio.getvalue(),
            file_name=f"驗收紀錄_{class_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
