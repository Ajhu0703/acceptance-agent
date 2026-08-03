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
    """超強效 PDF 解析：同時支援表格與純文字逐行掃描"""
    items = {}
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # 方式 1: 嘗試抽取表格
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row: continue
                    clean_row = [str(c).replace('\n', ' ').strip() if c else "" for c in row]
                    if clean_row[0].isdigit():
                        p_num = clean_row[0]
                        items[p_num] = {
                            "name": clean_row[1] if len(clean_row) > 1 else "",
                            "spec": clean_row[2] if len(clean_row) > 2 else "",
                            "qty": clean_row[4] if len(clean_row) > 4 else (clean_row[3] if len(clean_row) > 3 else "1")
                        }
            
            # 方式 2: 若表格抓不到，直接抽取純文字逐行掃描
            if not items:
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        parts = line.split()
                        # 如果該行開頭第一個字是數字 (品號)
                        if parts and parts[0].isdigit():
                            p_num = parts[0]
                            p_name = parts[1] if len(parts) > 1 else ""
                            p_spec = parts[2] if len(parts) > 2 else ""
                            p_qty = parts[-1] if len(parts) > 3 else "1"
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
        
        if items_dict:
            st.info(f"✅ PDF 解析成功！共抓取到 {len(items_dict)} 筆品號資料。")
        else:
            st.error("❌ PDF 解析失敗，可能是掃描圖片檔。請確認 PDF 檔是否能複製文字！")

        # B. 處理照片排序
        sorted_images = []
        if uploaded_images:
            def extract_img_number(img_file):
                nums = re.findall(r'\d+', img_file.name)
                return int(nums[-1]) if nums else 999
            sorted_images = sorted(uploaded_images, key=extract_img_number)

        # C. 載入 Word 範本
        doc = docx.Document(uploaded_doc)

        img_idx = 0
        # D. 逐一比對與替換
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()

                    # 1. 抬頭
                    if "驗收單號" in cell_text or "班級" in cell_text or "115W0149" in cell_text:
                        cell.text = class_name
                    elif "進貨日" in cell_text:
                        cell.text = f"進貨日：{date_str}"

                    # 2. 照片
                    elif "照片" in cell_text:
                        cell.text = ""
                        if img_idx < len(sorted_images):
                            p = cell.paragraphs[0]
                            p.add_run().add_picture(sorted_images[img_idx], width=Inches(2.2))
                            img_idx += 1

                    # 3. 品號文字格替換
                    elif "品號" in cell_text or "品名" in cell_text:
                        nums = re.findall(r'\d+', cell_text)
                        if nums:
                            p_num = nums[0]
                            if p_num in items_dict:
                                info = items_dict[p_num]
                                cell.text = f"品號{p_num} {info['name']} {info['spec']} 數量:{info['qty']}"
                            else:
                                cell.text = f"品號{p_num} (未於PDF找到品名)"

        # E. 加會驗結果
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
