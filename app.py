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
uploaded_images = st.file_uploader("3. 上傳材料照片 (可多選，按品號順序)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

def parse_pdf_items(pdf_file):
    """解析 PDF 提取品號、品名、規格、數量"""
    items = {}
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # 濾除空行或表頭
                    if not row or "品號" in str(row[0]):
                        continue
                    # 處理欄位資料
                    try:
                        p_no = str(row[0]).strip()
                        p_name = str(row[1]).strip() if len(row) > 1 else ""
                        p_spec = str(row[2]).strip() if len(row) > 2 else ""
                        p_qty = str(row[4]).strip() if len(row) > 4 else ""
                        
                        if p_no.isdigit():
                            items[p_no] = {
                                "name": p_name,
                                "spec": p_spec,
                                "qty": p_qty
                            }
                    except Exception:
                        continue
    return items

# 3. 產生文件按鈕
if st.button("🚀 產生驗收單文件"):
    if not uploaded_doc:
        st.error("請上傳空白 Word 範本 (.docx)！")
    elif not uploaded_pdf:
        st.error("請上傳材料明細 PDF (.pdf)！")
    else:
        # A. 解析 PDF 材料清單
        items_dict = parse_pdf_items(uploaded_pdf)
        
        # B. 載入 Word 範本
        doc = docx.Document(uploaded_doc)
        
        # 替換抬頭資訊
        for p in doc.paragraphs:
            if "班級" in p.text or "驗收單號" in p.text:
                p.text = f"班級名稱：{class_name}"
            if "進貨日" in p.text:
                p.text = f"進貨日期：{date_str}"

        # 圖片列表與計數器
        img_list = list(uploaded_images) if uploaded_images else []
        img_idx = 0

        # C. 遍歷表格，填入文字與照片
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # 填寫頂端表頭
                    if "驗收單號" in cell.text:
                        cell.text = class_name
                    elif "進貨日" in cell.text:
                        cell.text = f"進貨日：{date_str}"
                    
                    # 填寫照片與資料
                    text = cell.text.strip()
                    
                    # 判斷是否為照片欄位 (如: 品號1的照片)
                    if "的照片" in text or "照片" in text:
                        if img_idx < len(img_list):
                            cell.text = "" # 清空原文字
                            p = cell.paragraphs[0]
                            p.add_run().add_picture(img_list[img_idx], width=Inches(2.2))
                            img_idx += 1
                            
                    # 判斷是否為文字欄位 (如: 品號1 品名 規格 數量)
                    match = re.search(r'品號\s*(\d+)', text)
                    if match:
                        p_num = match.group(1)
                        if p_num in items_dict:
                            info = items_dict[p_num]
                            cell.text = f"品號{p_num} {info['name']} {info['spec']} 數量：{info['qty']}"

        # D. 末端加上會驗結果聲明
        p = doc.add_paragraph()
        p.add_run("會驗結果：經確認，到貨物品均於效期內，且與規格相符。").bold = True

        # E. 輸出下載
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("🎉 文件生成成功！已自動將 PDF 材料資料與照片帶入！")
        st.download_button(
            label="📥 下載完成的 Word 檔",
            data=bio.getvalue(),
            file_name=f"驗收紀錄_{class_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
