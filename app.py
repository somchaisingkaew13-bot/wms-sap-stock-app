import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="WMS & SAP Stock Reconciliation V10", page_icon="📦", layout="wide")

st.title("📦 WMS & SAP Stock Reconciliation V10 (Web Version)")
st.markdown("**โปรแกรมเปรียบเทียบ Stock จากไฟล์ Export ของ SAP และ WMS (รองรับ HOLD Alerts และ UOM Mismatch)**")

st.info("""
**เงื่อนไขการประมวลผล:**
- **ยอด WMS**: กรองรหัสสินค้าที่ขึ้นต้นด้วย 3, 4 ออก และแยกสถานะ HOLD ออกไปแจ้งเตือนแยก (ไม่นำมารวมเทียบ Stock)
- **ยอด SAP**: นำ Unrestricted + Quality Inspection + Blocked มารวมกันเป็นยอด Total SAP Stock
- **Transit & Transfer**: แยกออกไปแสดงในชีตเฉพาะ (ไม่นำมารวมเทียบ Stock)
- **ตรวจสอบหน่วยนับ (UOM)**: หากหน่วยไม่ตรงกัน จะระบุในช่องหมายเหตุพร้อมแสดงหน่วยที่ต่างกัน
""")

col1, col2 = st.columns(2)
with col1:
    sap_file = st.file_uploader("1. อัปโหลดไฟล์ SAP (.xlsx, .xls)", type=['xlsx', 'xls'])
with col2:
    wms_file = st.file_uploader("2. อัปโหลดไฟล์ WMS (.xlsx, .xls)", type=['xlsx', 'xls'])

def process_data(sap_f, wms_f):
    # 1. อ่านและประมวลผล SAP
    df_sap = pd.read_excel(sap_f, sheet_name=0)
    df_sap['Material_Clean'] = df_sap['Material'].astype(str).str.strip().apply(
        lambda x: str(int(float(x))) if x.replace('.', '', 1).isdigit() else x
    )

    sap_grouped = df_sap.groupby('Material_Clean', as_index=False).agg(
        SAP_Description=('Material description', 'first'),
        SAP_Unit=('Base Unit of Measure', 'first'),
        SAP_Unrestricted=('Unrestricted', 'sum'),
        SAP_Quality=('Quality Inspection', 'sum'),
        SAP_Blocked=('Blocked', 'sum'),
        SAP_Transit=('Transit and Transfer', 'sum')
    )
    sap_grouped['SAP_Total_Stock'] = sap_grouped['SAP_Unrestricted'] + sap_grouped['SAP_Quality'] + sap_grouped['SAP_Blocked']
    df_transit = sap_grouped[sap_grouped['SAP_Transit'] > 0].copy()

    # 2. อ่านและประมวลผล WMS
    df_wms = pd.read_excel(wms_f, sheet_name='WM Landscape', header=2)
    df_wms = df_wms.dropna(subset=['Item']).copy()
    df_wms['Item_Clean'] = df_wms['Item'].astype(str).str.strip().apply(
        lambda x: str(int(float(x))) if x.replace('.', '', 1).isdigit() else x
    )

    df_wms_filtered = df_wms[~df_wms['Item_Clean'].str.startswith(('3', '4'))].copy()
    df_wms_hold = df_wms_filtered[df_wms_filtered['Status'] == 'HOLD'].copy()
    df_wms_valid = df_wms_filtered[df_wms_filtered['Status'] != 'HOLD'].copy()

    wms_grouped = df_wms_valid.groupby('Item_Clean', as_index=False).agg(
        WMS_Description=('Description', 'first'),
        WMS_Unit=('UOM', 'first'),
        WMS_Total_Qty=('Qty', 'sum')
    )

    # 3. รวมข้อมูลเพื่อเปรียบเทียบ
    merged = pd.merge(wms_grouped, sap_grouped, left_on='Item_Clean', right_on='Material_Clean', how='outer')
    merged['Item_Code'] = merged['Item_Clean'].combine_first(merged['Material_Clean'])
    merged['Description'] = merged['WMS_Description'].combine_first(merged['SAP_Description'])
    
    merged['WMS_Unit_Clean'] = merged['WMS_Unit'].fillna('-').astype(str).str.strip()
    merged['SAP_Unit_Clean'] = merged['SAP_Unit'].fillna('-').astype(str).str.strip()

    merged['WMS_Total_Qty'] = merged['WMS_Total_Qty'].fillna(0)
    merged['SAP_Unrestricted'] = merged['SAP_Unrestricted'].fillna(0)
    merged['SAP_Quality'] = merged['SAP_Quality'].fillna(0)
    merged['SAP_Blocked'] = merged['SAP_Blocked'].fillna(0)
    merged['SAP_Transit'] = merged['SAP_Transit'].fillna(0)
    merged['SAP_Total_Stock'] = merged['SAP_Total_Stock'].fillna(0)

    merged['Diff_Vs_Total_SAP'] = merged['WMS_Total_Qty'] - merged['SAP_Total_Stock']

    def classify_status(row):
        diff = round(row['Diff_Vs_Total_SAP'], 4)
        if diff == 0: return 'ตรงกัน'
        elif diff > 0: return 'WMS สูงกว่า SAP'
        else: return 'WMS ต่ำกว่า SAP'

    merged['Status'] = merged.apply(classify_status, axis=1)

    def classify_remarks(row):
        w_uom = row['WMS_Unit_Clean']
        s_uom = row['SAP_Unit_Clean']
        remarks = []
        if w_uom != '-' and s_uom != '-' and w_uom.upper() != s_uom.upper():
            remarks.append(f"หน่วยใน SAP ไม่ตรงกับ WMS (SAP: {s_uom}, WMS: {w_uom})")
        elif w_uom == '-' and s_uom != '-':
            remarks.append(f"ไม่มีข้อมูลใน WMS (หน่วย SAP: {s_uom})")
        elif s_uom == '-' and w_uom != '-':
            remarks.append(f"ไม่มีข้อมูลใน SAP (หน่วย WMS: {w_uom})")
        return " | ".join(remarks) if remarks else "-"

    merged['หมายเหตุ'] = merged.apply(classify_remarks, axis=1)
    merged = merged.sort_values(by='Item_Code')

    # 4. สร้างไฟล์ Excel 6 ชีตสมบูรณ์
    output = io.BytesIO()
    wb = openpyxl.Workbook()

    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    orange_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    red_font = Font(name="Segoe UI", size=10, color="C00000", bold=True)
    orange_font = Font(name="Segoe UI", size=10, color="B25900", bold=True)
    green_font = Font(name="Segoe UI", size=10, color="385723")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

    # Sheet 1: สรุปภาพรวม
    ws_summary = wb.active
    ws_summary.title = "สรุปภาพรวม"
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary.merge_cells("A1:D1")
    ws_summary["A1"] = "สรุปรายงานเปรียบเทียบ Stock V10 (Web App)"
    ws_summary["A1"].font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    ws_summary["A1"].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws_summary["A1"].alignment = Alignment(horizontal="center", vertical="center")

    uom_diff_count = len(merged[merged['หมายเหตุ'].str.contains('หน่วยใน SAP ไม่ตรงกับ WMS', na=False)])

    kpis = [
        ("จำนวนรายการสินค้าใน SAP ทั้งหมด", len(sap_grouped)),
        ("จำนวนรายการ WMS ทั้งหมด (ก่อนกรอง)", len(df_wms)),
        ("จำนวนรายการ WMS ที่กรองออก (รหัสขึ้นต้นด้วย 3, 4)", len(df_wms) - len(df_wms_filtered)),
        ("จำนวนรายการ WMS HOLD (แยกแจ้ง ไม่รวมเปรียบเทียบ)", len(df_wms_hold)),
        ("จำนวนรายการ WMS ปกติที่นำมาเปรียบเทียบ", len(wms_grouped)),
        ("จำนวนรายการสินค้าทั้งหมดที่เปรียบเทียบ", len(merged)),
        ("จำนวนรายการที่ Stock ตรงกัน", len(merged[merged['Status'] == 'ตรงกัน'])),
        ("จำนวนรายการที่ Stock ไม่ตรงกัน (Stock Discrepancy)", len(merged[merged['Status'] != 'ตรงกัน'])),
        ("จำนวนรายการที่หน่วยนับไม่ตรงกัน (UOM Mismatch)", uom_diff_count),
        ("จำนวนรายการสินค้า Transit & Transfer (แยกแจ้ง)", len(df_transit))
    ]

    ws_summary.cell(row=3, column=1, value="รายการสรุป").font = Font(name="Segoe UI", bold=True, color="FFFFFF")
    ws_summary.cell(row=3, column=2, value="จำนวน (รายการ)").font = Font(name="Segoe UI", bold=True, color="FFFFFF")
    for c in [1, 2]:
        ws_summary.cell(row=3, column=c).fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")

    r = 4
    for label, val in kpis:
        ws_summary.cell(row=r, column=1, value=label).font = Font(name="Segoe UI", size=11)
        c_val = ws_summary.cell(row=r, column=2, value=val)
        c_val.font = Font(name="Segoe UI", size=11, bold=True)
        c_val.number_format = '#,##0'
        r += 1

    headers = [
        'รหัสสินค้า (Material)', 'รายละเอียดสินค้า (Description)', 'หน่วย WMS', 'หน่วย SAP',
        'ยอดรวม WMS ปกติ (Qty)', 'SAP Unrestricted', 'SAP Quality Inspection', 'SAP Blocked', 'ยอดรวม SAP Stock (Unrestricted+Quality+Blocked)',
        'ผลต่าง Stock (WMS - Total SAP)', 'สถานะ Stock', 'หมายเหตุ (UOM Check)'
    ]

    def populate_sheet(ws, dataframe, title_bg="1F4E78"):
        ws.views.sheetView[0].showGridLines = True
        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = PatternFill(start_color=title_bg, end_color=title_bg, fill_type="solid")
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for _, row in dataframe.iterrows():
            row_vals = [
                row['Item_Code'], row['Description'], row['WMS_Unit_Clean'], row['SAP_Unit_Clean'],
                row['WMS_Total_Qty'], row['SAP_Unrestricted'], row['SAP_Quality'], row['SAP_Blocked'], row['SAP_Total_Stock'],
                row['Diff_Vs_Total_SAP'], row['Status'], row['หมายเหตุ']
            ]
            ws.append(row_vals)
            curr_row = ws.max_row
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=curr_row, column=c)
                cell.font = Font(name="Segoe UI", size=10)
                cell.border = thin_border
                
                if c in [5, 6, 7, 8, 9, 10]:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right")
                elif c in [1, 3, 4, 11]:
                    cell.alignment = Alignment(horizontal="center")
                
                if c == 11:
                    if row['Status'] != 'ตรงกัน':
                        cell.fill = red_fill
                        cell.font = red_font
                    else:
                        cell.font = green_font
                
                if c == 12:
                    if 'หน่วยใน SAP ไม่ตรงกับ WMS' in str(row['หมายเหตุ']):
                        cell.fill = orange_fill
                        cell.font = orange_font

    # Sheet 2: ข้อมูลเปรียบเทียบทั้งหมด
    ws_all = wb.create_sheet(title="ข้อมูลเปรียบเทียบทั้งหมด")
    populate_sheet(ws_all, merged, title_bg="1F4E78")

    # Sheet 3: เฉพาะ Stock ไม่ตรงกัน
    ws_diff = wb.create_sheet(title="เฉพาะ Stock ไม่ตรงกัน")
    populate_sheet(ws_diff, merged[merged['Status'] != 'ตรงกัน'], title_bg="C00000")

    # Sheet 4: เฉพาะหน่วยนับไม่ตรงกัน
    ws_uom = wb.create_sheet(title="เฉพาะหน่วยนับไม่ตรงกัน")
    populate_sheet(ws_uom, merged[merged['หมายเหตุ'].str.contains('หน่วยใน SAP ไม่ตรงกับ WMS', na=False)], title_bg="C65911")

    # Sheet 5: รายการ HOLD Alerts
    ws_hold = wb.create_sheet(title="รายการ HOLD (Alerts)")
    ws_hold.views.sheetView[0].showGridLines = True
    hold_headers = ['รหัสสินค้า (Item)', 'รายละเอียดสินค้า (Description)', 'LPN', 'Location', 'Batch SAP', 'ยอด Qty HOLD', 'หน่วย (UOM)', 'สถานะ WMS', 'หมายเหตุ']
    ws_hold.append(hold_headers)
    for col_idx in range(1, len(hold_headers) + 1):
        cell = ws_hold.cell(row=1, column=col_idx)
        cell.fill = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if len(df_wms_hold) > 0:
        for _, row in df_wms_hold.iterrows():
            ws_hold.append([
                str(row['Item_Clean']), row['Description'], row['LPN'], row['Location'],
                row['Batch SAP'], row['Qty'], row['UOM'], row['Status'], 'รายการ HOLD ไม่ถูกนำมารวมในยอด WMS เปรียบเทียบ'
            ])
            curr_row = ws_hold.max_row
            for c in range(1, len(hold_headers) + 1):
                cell = ws_hold.cell(row=curr_row, column=c)
                cell.font = Font(name="Segoe UI", size=10)
                cell.border = thin_border
                if c == 6:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right")
                elif c in [1, 3, 4, 7, 8]:
                    cell.alignment = Alignment(horizontal="center")
                if c == 8:
                    cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")

    # Sheet 6: สินค้าระหว่างทาง (Transit)
    ws_transit = wb.create_sheet(title="สินค้าระหว่างทาง (Transit)")
    ws_transit.views.sheetView[0].showGridLines = True
    transit_headers = ['รหัสสินค้า (Material)', 'รายละเอียดสินค้า (Description)', 'หน่วยนับ (Unit)', 'จำนวน Transit & Transfer ใน SAP']
    ws_transit.append(transit_headers)
    for col_idx in range(1, len(transit_headers) + 1):
        cell = ws_transit.cell(row=1, column=col_idx)
        cell.fill = PatternFill(start_color="333F48", end_color="333F48", fill_type="solid")
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if len(df_transit) > 0:
        for _, row in df_transit.iterrows():
            ws_transit.append([row['Material_Clean'], row['SAP_Description'], row['SAP_Unit'], row['SAP_Transit']])
            curr_row = ws_transit.max_row
            for c in range(1, len(transit_headers) + 1):
                cell = ws_transit.cell(row=curr_row, column=c)
                cell.font = Font(name="Segoe UI", size=10)
                cell.border = thin_border
                if c == 4:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right")
                elif c in [1, 3]:
                    cell.alignment = Alignment(horizontal="center")
    else:
        ws_transit.append(["-", "ไม่พบรายการสินค้าระหว่างทาง (Transit & Transfer) ในไฟล์ SAP นี้", "-", 0])

    # ปรับความกว้างคอลัมน์อัตโนมัติ
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                max_len = max(max_len, len(val_str))
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output)
    output.seek(0)
    return output, kpis

if sap_file is not None and wms_file is not None:
    if st.button("▶ เริ่มเปรียบเทียบ Stock", type="primary", use_container_width=True):
        with st.spinner('กำลังประมวลผลเปรียบเทียบข้อมูล 6 ชีตสมบูรณ์...'):
            try:
                output_file, summary_kpis = process_data(sap_file, wms_file)
                st.success("✅ ประมวลผลสำเร็จ! (สร้างรายงาน 6 ชีตครบถ้วน)")
                
                st.subheader("📊 สรุปผลลัพธ์เบื้องต้น")
                for label, val in summary_kpis:
                    st.write(f"- **{label}**: {val:,} รายการ")

                st.download_button(
                    label="📥 ดาวน์โหลดรายงานผลลัพธ์ Excel ฉบับสมบูรณ์ V10 (.xlsx)",
                    data=output_file,
                    file_name="Stock_Comparison_Report_V10.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผล: {e}")
else:
    st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง SAP และ WMS ด้านบน")
