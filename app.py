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
    """強力版 PDF 文字解析：直接逐行比對提取品號、品名、規格、數量"""
    items = {}
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                parts = line.split()
                # 當一行開頭是數字（品號），且包含品名相關資料時
                if parts and parts[0].isdigit():
                    p_num = parts[0]
                    # 抓取該品號對應的剩餘欄位
                    p_name = parts[1] if len(parts) > 1 else ""
                    p_spec = parts[2] if len(parts) > 2 else ""
                    # 尋找數量 (單位後面的數字，如 包 1 -> 1)
                    p_qty = parts[4] if len(parts) > 4 else (parts[3] if len(parts) > 3 else "1")
                    
                    items[p_num] = {
                        "name": p_name,
                        "spec": p_spec,
                        "qty": p_qty
                    }
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

        # 圖片列表
        img_list = list(uploaded_images) if uploaded_images else []
        img_idx = 0

        # C. 遍歷 Word 表格，替換文字與插入照片
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.strip()

                    # 1. 填寫頂端抬頭
                    if "驗收單號" in cell_text or "班級" in cell_text:
                        cell.text = class_name
                    elif "進貨日" in cell_text:
                        cell.text = f"進貨日：{date_str}"

                    # 2. 插入照片
                    elif "照片" in cell_text or "的照片" in cell_text:
                        if img_idx < len(img_list):
                            cell.text = "" # 清空原標籤
                            p = cell.paragraphs[0]
                            p.add_run().add_picture(img_list[img_idx], width=Inches(2.2))
                            img_idx += 1

                    # 3. 填寫品號文字 (搜尋 品號1, 品號2... 或 品號 1)
                    elif "品號" in cell_text:
                        match = re.search(r'品號\s*(\d+)', cell_text)
                        if match:
                            p_num = match.group(1)
                            if p_num in items_dict:
                                info = items_dict[p_num]
                                cell.text = f"品號{p_num} {info['name']} {info['spec']} 數量:{info['qty']}"
                            else:
                                cell.text = f"品號{p_num}"

        # D. 末端加上會驗結果聲明
        p = doc.add_paragraph()
        p.add_run("會驗結果：經確認，到貨物品均於效期內，且與規格相符。").bold = True

        # E. 輸出下載
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("🎉 文件生成成功！已成功將 PDF 品號明細與照片同步帶入！")
        st.download_button(
            label="📥 下載完成的 Word 檔",
            data=bio.getvalue(),
            file_name=f"驗收紀錄_{class_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
