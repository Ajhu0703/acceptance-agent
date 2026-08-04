import streamlit as st
import docx
from docx.shared import Inches
import pdfplumber
import io
import re

st.set_page_config(page_title="材料驗收單自動生成器", layout="centered")

st.title("📋 材料驗收單自動生成器")

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
    """修正欄位對應：精準抓取品號、品名、規格、數量 (過濾單價與金額)"""
    items = {}
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # 1. 優先使用表格解析
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row: continue
                    clean_row = [str(c).replace('\n', ' ').strip() if c else "" for c in row]
                    # 當第一欄是數字（品號）
                    if clean_row[0].isdigit():
                        p_num = clean_row[0]
                        p_name = clean_row[1] if len(clean_row) > 1 else ""
                        p_spec = clean_row[2] if len(clean_row) > 2 else ""
                        
                        # 修正數量：判斷欄位長度，通常數量在倒數第 2 欄或第 5 欄 (index 4 或 5)
                        # 如果有 6 個欄位 (品號, 品名, 規格, 單價, 數量, 金額)，數量為 index 4
                        p_qty = "1"
                        if len(clean_row) >= 6:
                            p_qty = clean_row[4]
                        elif len(clean_row) == 5:
                            p_qty = clean_row[3]
                        
                        items[p_num] = {
                            "name": p_name,
                            "spec": p_spec,
                            "qty": p_qty
                        }
            
            # 2. 備用純文字解析
            if not items:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        parts = line.split()
                        if parts and parts[0].isdigit():
                            p_num = parts[0]
                            p_name = parts[1] if len(parts) > 1 else ""
                            p_spec = parts[2] if len(parts) > 2 else ""
                            # 純文字情況下，數量通常在規格後的數字
                            p_qty = parts[4] if len(parts) > 5 else (parts[3] if len(parts) > 4 else "1")
                            items[p_num] = {
                                "name": p_name,
                                "spec": p_spec,
                                "qty": p_qty
                            }
    return items

if st.button("🚀 產生驗收單文件"):
    if not uploaded_doc or not uploaded_pdf:
        st.error("請確保上傳了 Word 範本與 PDF 明細檔！")
    else:
        # A. 解析 PDF
        items_dict = parse_pdf_items(uploaded_pdf)
        
        # B. 精準處理照片檔名括號內的數字：如 U-2026-08-03(1).jpg -> 抓出 1
        image_map = {}
        if uploaded_images:
            for img in uploaded_images:
                # 優先尋找括號內的數字 (1) 或檔名最後面的數字
                match = re.search(r'\((\d+)\)', img.name)
                if match:
                    p_num = match.group(1)
                else:
                    nums = re.findall(r'\d+', img.name)
                    p_num = nums[-1] if nums else None
                
                if p_num:
                    image_map[p_num] = img

        # C. 載入 Word 範本
        doc = docx.Document(uploaded_doc)

        # D. 遍歷 Word 表格與替換
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()

                    # 1. 抬頭
                    if "驗收單號" in cell_text or "班級" in cell_text or "115W0149" in cell_text:
                        cell.text = class_name
                    elif "進貨日" in cell_text:
                        cell.text = f"進貨日：{date_str}"

                    # 2. 品號與照片比對 (以品號數字發動)
                    elif "品號" in cell_text or "品名" in cell_text:
                        nums = re.findall(r'\d+', cell_text)
                        if nums:
                            p_num = nums[0]
                            
                            # 如果這格是照片格
                            if "照片" in cell_text or "的照片" in cell_text:
                                cell.text = ""
                                if p_num in image_map:
                                    p = cell.paragraphs[0]
                                    p.add_run().add_picture(image_map[p_num], width=Inches(2.2))
                            
                            # 如果這格是文字說明格
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
        
        st.success("🎉 驗收單文件生成成功！數量與照片對應已更新！")
        st.download_button(
            label="📥 下載完成的 Word 檔",
            data=bio.getvalue(),
            file_name=f"驗收紀錄_{class_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
