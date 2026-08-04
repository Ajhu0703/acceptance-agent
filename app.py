import streamlit as st
import docx
from docx.shared import Inches
import pdfplumber
import io
import re

st.set_page_config(page_title="材料驗收單自動生成器", layout="centered")

st.title("📋 材料驗收單自動生成器")

# 1. 基本資訊填寫區
col1, col2 = st.columns(2)
with col1:
    class_name = st.text_input("班級名稱", value="115W0149-飲調輕食暨餐飲創業(桃園)-第1期")
with col2:
    date_str = st.text_input("進貨日期", value="115/02/03")

st.markdown("---")

# 2. 檔案上傳區
uploaded_pdf = st.file_uploader("1. 上傳材料明細 PDF (.pdf)", type=["pdf"])
uploaded_doc = st.file_uploader("2. 上傳空白 Word 範本 (.docx)", type=["docx"])
uploaded_images = st.file_uploader(
    "3. 上傳材料照片 (檔名請含品號數字，如 1.jpg 或 (1).jpg)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True
)

def parse_pdf_items(pdf_file):
    """精準解析 PDF：自動識別品號、品名、規格與數量 (過濾單價與金額)"""
    items = {}
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # 優先嘗試表格擷取
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row:
                        continue
                    clean_row = [str(c).replace('\n', ' ').strip() if c else "" for c in row]
                    # 判斷第一欄是否為數字 (品號)
                    if clean_row[0].isdigit():
                        p_num = clean_row[0]
                        p_name = clean_row[1] if len(clean_row) > 1 else ""
                        p_spec = clean_row[2] if len(clean_row) > 2 else ""
                        
                        # 數量智慧判斷：根據欄位數量過濾單價
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
            
            # 若表格未擷取成功，使用備用純文字逐行掃描
            if not items:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        parts = line.split()
                        if parts and parts[0].isdigit():
                            p_num = parts[0]
                            p_name = parts[1] if len(parts) > 1 else ""
                            p_spec = parts[2] if len(parts) > 2 else ""
                            p_qty = parts[4] if len(parts) > 5 else (parts[3] if len(parts) > 4 else "1")
                            items[p_num] = {
                                "name": p_name,
                                "spec": p_spec,
                                "qty": p_qty
                            }
    return items

# 3. 文件生成邏輯
if st.button("🚀 產生驗收單文件"):
    if not uploaded_doc or not uploaded_pdf:
        st.error("請確保上傳了 Word 範本與 PDF 明細檔！")
    else:
        # A. 解析 PDF 材料清單
        items_dict = parse_pdf_items(uploaded_pdf)
        
        # B. 建立照片與品號的印射字典
        image_map = {}
        if uploaded_images:
            for img in uploaded_images:
                # 優先抓取括號內的數字，如 (1).jpg -> 1；否則抓取檔名最後一組數字
                match = re.search(r'\((\d+)\)', img.name)
                if match:
                    p_num = match.group(1)
                else:
                    nums = re.findall(r'\d+', img.name)
                    p_num = nums[-1] if nums else None
                
                if p_num:
                    image_map[p_num] = img

        # C. 載入 Word 範本並替換內容
        doc = docx.Document(uploaded_doc)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()

                    # 1. 填寫標頭資訊
                    if "驗收單號" in cell_text or "班級" in cell_text or "115W0149" in cell_text:
                        cell.text = class_name
                    elif "進貨日" in cell_text:
                        cell.text = f"進貨日：{date_str}"

                    # 2. 處理表格內的照片格與文字格
                    elif "品號" in cell_text or "品名" in cell_text:
                        nums = re.findall(r'\d+', cell_text)
                        if nums:
                            p_num = nums[0]
                            
                            # 處理照片插入
                            if "照片" in cell_text or "的照片" in cell_text:
                                cell.text = "" # 清空原文字
                                if p_num in image_map:
                                    p = cell.paragraphs[0]
                                    p.add_run().add_picture(image_map[p_num], width=Inches(2.2))
                            
                            # 處理文字資料填寫
                            else:
                                if p_num in items_dict:
                                    info = items_dict[p_num]
                                    cell.text = f"品號{p_num} {info['name']} {info['spec']} 數量:{info['qty']}"

        # D. 新增末端會驗結果
        p = doc.add_paragraph()
        p.add_run("會驗結果：經確認，到貨物品均於效期內，且與規格相符。").bold = True

        # E. 輸出與下載處理
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("🎉 驗收單文件生成成功！照片與資料已精準填入。")
        st.download_button(
            label="📥 下載完成的 Word 檔",
            data=bio.getvalue(),
            file_name=f"驗收紀錄_{class_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
