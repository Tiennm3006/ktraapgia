import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime

# Function to create Word report
def create_word_report(df_summary, top_3, bottom_3, chart_buf, top3_chart_buf, bottom3_chart_buf, pie_chart_buf, total_checks):
    doc = Document()

    doc.add_heading('Báo Cáo Đánh Giá Công Tác Kiểm Tra Áp Giá', 0)

    # Add Summary
    doc.add_heading('1. Tổng quan:', level=1)
    doc.add_paragraph(f'Tổng số lượng kiểm tra toàn công ty: {total_checks:,} lượt.')
    if total_checks == 0:
        doc.add_paragraph("Chưa có dữ liệu kiểm tra.")
    else:
        doc.add_paragraph("Báo cáo tập trung vào tỷ lệ số khách hàng có thay đổi sau kiểm tra trên tổng số lượng kiểm tra. Các điện lực có tỷ lệ cao nhất và thấp nhất được minh họa qua biểu đồ sau.")

    # Đổi tên cột
    column_mapping = {
        'Area': 'Điện lực',
        'Total_Checks': 'Tổng số lượng kiểm tra',
        'Total_Changes': 'Số KH có thay đổi'
    }
    df_export = df_summary.rename(columns=column_mapping)

    # Add data table
    doc.add_heading('2. Bảng tổng hợp:', level=1)
    table = doc.add_table(rows=1, cols=len(df_export.columns))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df_export.columns):
        hdr_cells[i].text = column

    for index, row in df_export.iterrows():
        if row['Điện lực'] and str(row['Điện lực']).strip().lower() not in ['nan', 'đơn vị']:
            row_cells = table.add_row().cells
            for i, item in enumerate(row):
                row_cells[i].text = str(item)

    # Add Top 3 and Bottom 3 analysis
    doc.add_heading('3. Điện lực tỷ lệ cao nhất:', level=1)
    for idx, row in top_3.iterrows():
        doc.add_paragraph(f"- {row['Area']}: {row['Tỷ lệ thay đổi (%)']}% ({row['Total_Checks']} lượt, {row['Total_Changes']} thay đổi)")

    doc.add_heading('4. Điện lực tỷ lệ thấp nhất:', level=1)
    for idx, row in bottom_3.iterrows():
        doc.add_paragraph(f"- {row['Area']}: {row['Tỷ lệ thay đổi (%)']}% ({row['Total_Checks']} lượt, {row['Total_Changes']} thay đổi)")

    # Add Charts
    doc.add_heading('5. Biểu Đồ Tỷ Lệ thay đổi:', level=1)
    doc.add_picture(chart_buf, width=Inches(6))

    doc.add_heading('6. Biểu Đồ Top 3:', level=1)
    doc.add_picture(top3_chart_buf, width=Inches(5))

    doc.add_heading('7. Biểu Đồ Bottom 3:', level=1)
    doc.add_picture(bottom3_chart_buf, width=Inches(5))

    doc.add_heading('8. Biểu Đồ Tỷ Lệ thay đổi vs không thay đổi:', level=1)
    doc.add_picture(pie_chart_buf, width=Inches(4.5))

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# MAIN APP RUN

st.set_page_config(page_title="Đánh giá kiểm tra áp giá", layout="wide")
st.title("Đánh giá công tác kiểm tra áp giá")

uploaded_file = st.file_uploader("Tải lên file Excel dữ liệu", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()
    df.columns = ['Stt', 'Area', 'Check_SH_2plus', 'Check_HCSN', 'Check_Production', 'Check_KDDV', 'Check_PriceRate', 'Check_SH_Level3', 'Total_Checks', 'Total_Changes']

    df_summary = df[['Area', 'Total_Checks', 'Total_Changes']].copy()
    df_summary['Total_Checks'] = pd.to_numeric(df_summary['Total_Checks'], errors='coerce')
    df_summary['Total_Changes'] = pd.to_numeric(df_summary['Total_Changes'], errors='coerce')
    df_summary = df_summary[df_summary['Area'].notna()]
    df_summary['Area'] = df_summary['Area'].astype(str)
    df_summary = df_summary[df_summary['Area'].str.strip().str.lower() != 'tổng cộng']

    total_checks = df_summary['Total_Checks'].sum()
    total_changes = df_summary['Total_Changes'].sum()
    df_summary['Tỷ lệ thay đổi (%)'] = (df_summary['Total_Changes'] / df_summary['Total_Checks'] * 100).round(2)

    top_3 = df_summary.sort_values(by='Tỷ lệ thay đổi (%)', ascending=False).head(3)
    bottom_3 = df_summary.sort_values(by='Tỷ lệ thay đổi (%)', ascending=True).head(3)

    fig, ax = plt.subplots(figsize=(12,6))
    bars = ax.bar(df_summary['Area'], df_summary['Tỷ lệ thay đổi (%)'])
    ax.set_title('Tỷ lệ thay đổi sau kiểm tra theo Điện lực')
    ax.set_ylabel('Tỷ lệ (%)')
    plt.xticks(rotation=90)
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{height:.2f}%', (bar.get_x() + bar.get_width() / 2, height), textcoords="offset points", xytext=(0,3), ha='center')
    chart_buf = BytesIO()
    fig.savefig(chart_buf)
    chart_buf.seek(0)
    st.pyplot(fig)

    # Top 3
    fig_top3, ax_top3 = plt.subplots()
    bars_top3 = ax_top3.bar(top_3['Area'], top_3['Tỷ lệ thay đổi (%)'], color='green')
    ax_top3.set_title("Top 3 Điện lực có tỷ lệ thay đổi cao nhất")
    for bar in bars_top3:
        height = bar.get_height()
        ax_top3.annotate(f'{height:.2f}%', (bar.get_x() + bar.get_width() / 2, height), textcoords="offset points", xytext=(0,3), ha='center')
    top3_chart_buf = BytesIO()
    fig_top3.savefig(top3_chart_buf)
    top3_chart_buf.seek(0)
    st.pyplot(fig_top3)

    # Bottom 3
    fig_bottom3, ax_bottom3 = plt.subplots()
    bars_bottom3 = ax_bottom3.bar(bottom_3['Area'], bottom_3['Tỷ lệ thay đổi (%)'], color='red')
    ax_bottom3.set_title("Bottom 3 Điện lực có tỷ lệ thay đổi thấp nhất")
    for bar in bars_bottom3:
        height = bar.get_height()
        ax_bottom3.annotate(f'{height:.2f}%', (bar.get_x() + bar.get_width() / 2, height), textcoords="offset points", xytext=(0,3), ha='center')
    bottom3_chart_buf = BytesIO()
    fig_bottom3.savefig(bottom3_chart_buf)
    bottom3_chart_buf.seek(0)
    st.pyplot(fig_bottom3)

    # Pie chart
    fig_pie, ax_pie = plt.subplots()
    ax_pie.pie([total_changes, total_checks - total_changes], labels=['Có thay đổi', 'Không thay đổi'], autopct='%1.1f%%', colors=['#ff9999','#66b3ff'])
    ax_pie.set_title('Tỷ lệ KH có thay đổi vs không thay đổi')
    pie_chart_buf = BytesIO()
    fig_pie.savefig(pie_chart_buf)
    pie_chart_buf.seek(0)
    st.pyplot(fig_pie)

    st.dataframe(df_summary)

    today = datetime.today().strftime('%Y-%m-%d')
    word_file = create_word_report(df_summary, top_3, bottom_3, chart_buf, top3_chart_buf, bottom3_chart_buf, pie_chart_buf, total_checks)

    st.download_button(
        label="📄 Tải báo cáo Word",
        data=word_file,
        file_name=f'Bao_cao_kiem_tra_ap_gia_{today}.docx',
        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
