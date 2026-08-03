import streamlit as st
import docx
from docx.shared import Inches
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

# 2. 檔案上傳
uploaded_doc = st.file_uploader("上傳空白 Word 範本 (.docx)", type=["docx"])
uploaded_images = st.file_uploader("上傳材料照片 (可多選)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# 3. 產生文件按鈕
if st.button("🚀 產生驗收單文件"):
    if not uploaded_doc:
        st.error("請先上傳空白 Word 範本！")
    else:
        # 載入 Word 範本
        doc = docx.Document(uploaded_doc)
        
        # 邏輯 A：替換文件段落中的頁首/抬頭資訊
        for p in doc.paragraphs:
            if "班級" in p.text or "驗收單號" in p.text:
                p.text = f"班級名稱：{class_name}"
            if "進貨日" in p.text:
                p.text = f"進貨日期：{date_str}"
        
        # 邏輯 B：如果有表格，替換表格內容與嵌入圖片
        img_list = list(uploaded_images) if uploaded_images else []
        img_idx = 0
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # 填寫抬頭
                    if "驗收單號" in cell.text:
                        cell.text = class_name
                    elif "進貨日" in cell.text:
                        cell.text = f"進貨日：{date_str}"
                    
                    # 自動插入照片到對應的照片格子
                    if "照片" in cell.text and img_idx < len(img_list):
                        cell.text = "" # 清空提示文字
                        p = cell.paragraphs[0]
                        p.add_run().add_picture(img_list[img_idx], width=Inches(2.2))
                        img_idx += 1
                        
        # 邏輯 C：末端自動加上會驗結果聲明
        p = doc.add_paragraph()
        p.add_run("會驗結果：經確認，到貨物品均於效期內，且與規格相符。").bold = True
        
        # 儲存至記憶體提供下載
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("🎉 文件生成成功！照片與欄位已自動填入！")
        st.download_button(
            label="📥 下載完成的 Word 檔",
            data=bio.getvalue(),
            file_name=f"驗收紀錄_{class_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
