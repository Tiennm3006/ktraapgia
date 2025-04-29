import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from docx import Document
from docx.shared import Inches
from datetime import datetime

# Function to create Word report
def create_word_report(df_summary, top_3, bottom_3, chart_buf, top3_chart_buf, bottom3_chart_buf, total_checks):
    doc = Document()

    doc.add_heading('Báo Cáo Đánh Giá Công Tác Kiểm Tra Áp Giá', 0)

    # Add Summary
    doc.add_heading('1. Tổng quan:', level=1)
    doc.add_paragraph(f'Tổng số lượng kiểm tra toàn công ty: {total_checks:,} lượt.')
    if total_checks == 0:
        doc.add_paragraph("Chưa có dữ liệu kiểm tra.")
    else:
        doc.add_paragraph("Công tác kiểm tra đã được thực hiện ở nhiều điện lực, trong đó nổi bật là các đơn vị có tỷ lệ thực hiện cao và những đơn vị còn hạn chế.")

    # Add data table
    doc.add_heading('2. Bảng tổng hợp:', level=1)
    table = doc.add_table(rows=1, cols=len(df_summary.columns))
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df_summary.columns):
        hdr_cells[i].text = column

    for index, row in df_summary.iterrows():
        if row['Area'] != 'nan':
            row_cells = table.add_row().cells
            for i, item in enumerate(row):
                row_cells[i].text = str(item)

    # Add Top 3 and Bottom 3 analysis
    doc.add_heading('3. Điện lực tỷ lệ cao nhất:', level=1)
    for idx, row in top_3.iterrows():
        doc.add_paragraph(f"- {row['Area']}: {row['Tỷ lệ (%)']}% ({row['Total_Checks']} lượt)")

    doc.add_heading('4. Điện lực tỷ lệ thấp nhất:', level=1)
    for idx, row in bottom_3.iterrows():
        doc.add_paragraph(f"- {row['Area']}: {row['Tỷ lệ (%)']}% ({row['Total_Checks']} lượt)")

    # Add Charts
    doc.add_heading('5. Biểu Đồ Tỷ Lệ Kiểm Tra:', level=1)
    doc.add_picture(chart_buf, width=Inches(6))

    doc.add_heading('6. Biểu Đồ Top 3 Điện lực cao nhất:', level=1)
    doc.add_picture(top3_chart_buf, width=Inches(5))

    doc.add_heading('7. Biểu Đồ Bottom 3 Điện lực thấp nhất:', level=1)
    doc.add_picture(bottom3_chart_buf, width=Inches(5))

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# Streamlit App
st.set_page_config(page_title="Đánh giá công tác kiểm tra áp giá", layout="wide")
st.title('Đánh giá công tác kiểm tra áp giá')

uploaded_file = st.file_uploader('Tải lên file Excel dữ liệu', type=['xlsx'])

if uploaded_file is not None:
    # Load data
    df = pd.read_excel(uploaded_file, sheet_name=0)

    # Chuẩn hoá tên cột
    df.columns = df.columns.str.strip()
    df.columns = ['Stt', 'Area', 'Check_SH_2plus', 'Check_HCSN',
                  'Check_Production', 'Check_KDDV', 'Check_PriceRate',
                  'Check_SH_Level3', 'Total_Checks', 'Total_Changes']
    df = df.reset_index(drop=True)

    df_summary = df[['Area', 'Total_Checks']].copy()
    df_summary['Total_Checks'] = pd.to_numeric(df_summary['Total_Checks'], errors='coerce')
    df_summary = df_summary[df_summary['Area'].notna()]
    df_summary['Area'] = df_summary['Area'].astype(str)

    # Bỏ 'Tổng cộng' nếu có
    df_summary = df_summary[df_summary['Area'].str.strip().str.lower() != 'tổng cộng']

    total_checks = df_summary['Total_Checks'].sum()
    df_summary['Tỷ lệ (%)'] = (df_summary['Total_Checks'] / total_checks * 100).round(2)

    # Top 3 & Bottom 3
    top_3 = df_summary.sort_values(by='Tỷ lệ (%)', ascending=False).head(3)
    bottom_3 = df_summary.sort_values(by='Tỷ lệ (%)', ascending=True).head(3)

    # Vẽ biểu đồ toàn bộ
    fig, ax = plt.subplots(figsize=(12,6))
    bars = ax.bar(df_summary['Area'], df_summary['Tỷ lệ (%)'])
    ax.set_title('Tỷ lệ kiểm tra theo Điện lực')
    ax.set_ylabel('Tỷ lệ (%)')
    plt.xticks(rotation=90)
    plt.grid(axis='y')
    plt.tight_layout()

    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{height:.2f}%',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=8)

    chart_buf = BytesIO()
    fig.savefig(chart_buf)
    chart_buf.seek(0)

    st.pyplot(fig)

    # Vẽ biểu đồ Top 3
    fig_top3, ax_top3 = plt.subplots(figsize=(8,5))
    bars_top3 = ax_top3.bar(top_3['Area'], top_3['Tỷ lệ (%)'], color='green')
    ax_top3.set_title('Top 3 Điện lực tỷ lệ cao nhất')
    ax_top3.set_ylabel('Tỷ lệ (%)')
    plt.xticks(rotation=45)
    plt.grid(axis='y')

    for bar in bars_top3:
        height = bar.get_height()
        ax_top3.annotate(f'{height:.2f}%',
                         xy=(bar.get_x() + bar.get_width() / 2, height),
                         xytext=(0, 3),
                         textcoords="offset points",
                         ha='center', va='bottom', fontsize=8)

    top3_chart_buf = BytesIO()
    fig_top3.savefig(top3_chart_buf)
    top3_chart_buf.seek(0)

    # Vẽ biểu đồ Bottom 3
    fig_bottom3, ax_bottom3 = plt.subplots(figsize=(8,5))
    bars_bottom3 = ax_bottom3.bar(bottom_3['Area'], bottom_3['Tỷ lệ (%)'], color='red')
    ax_bottom3.set_title('Bottom 3 Điện lực tỷ lệ thấp nhất')
    ax_bottom3.set_ylabel('Tỷ lệ (%)')
    plt.xticks(rotation=45)
    plt.grid(axis='y')

    for bar in bars_bottom3:
        height = bar.get_height()
        ax_bottom3.annotate(f'{height:.2f}%',
                            xy=(bar.get_x() + bar.get_width() / 2, height),
                            xytext=(0, 3),
                            textcoords="offset points",
                            ha='center', va='bottom', fontsize=8)

    bottom3_chart_buf = BytesIO()
    fig_bottom3.savefig(bottom3_chart_buf)
    bottom3_chart_buf.seek(0)

    st.pyplot(fig_top3)
    st.pyplot(fig_bottom3)

    st.dataframe(df_summary)

    # Export Word
    today = datetime.today().strftime('%Y-%m-%d')
    word_file = create_word_report(df_summary, top_3, bottom_3, chart_buf, top3_chart_buf, bottom3_chart_buf, total_checks)

    st.download_button(
        label="📄 Tải báo cáo Word",
        data=word_file,
        file_name=f'Bao_cao_kiem_tra_ap_gia_{today}.docx',
        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
