import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from io import BytesIO

st.set_page_config(page_title="Geometric Data Reader", layout="wide")

def parse_data(file_content):
    """Parse data from file content string"""
    if isinstance(file_content, bytes):
        lines = file_content.decode('utf-8').splitlines()
    else:
        lines = file_content.splitlines()
    
    data = []
    
    for line in lines:
        if not line.strip():
            continue
            
        parts = line.strip().split(';')
        if len(parts) < 2:
            continue
        
        obj_id = parts[0]
        obj_type = parts[1]
        row = {'ID': obj_id, 'Type': obj_type}
        remaining_values = [p.strip() for p in parts[2:]]
        
        if obj_type == 'PLANE':
            cols = ['Method', 'X', 'Y', 'Z', 'A', 'B', 'C', '', 'D', 'Dev']
        elif obj_type == 'CIRCLE':
            cols = ['Method', 'X', 'Y', 'Z', 'I', 'J', 'K', '', 'Radius', 'Dev']
        elif obj_type == 'PT-COMP':
            cols = ['Method', 'X', 'Y', 'Z']
        elif obj_type == 'DISTANCE':
            cols = ['', 'X', 'Y', 'Z', '', '', '', '', 'Distance']
        elif obj_type == 'CONE':
            cols = ['Method', 'X', 'Y', 'Z', 'I', 'J', 'K', '', 'Half-Angle', 'Dev']
        elif obj_type == 'INT-CIRCLE':
            cols = ['', 'X', 'Y', 'Z', 'I', 'J', 'K', '', 'Radius']
        elif obj_type == 'SYM-POINT':
            cols = ['', 'X', 'Y', 'Z']
        else:
            cols = [f'Val{i+1}' for i in range(len(remaining_values))]
        
        for i, val in enumerate(remaining_values):
            if i < len(cols) and cols[i] and val:
                numeric_pattern = r'^-?\d+\.?\d*$'
                if re.match(numeric_pattern, val):
                    row[cols[i]] = float(val)
                else:
                    row[cols[i]] = val
        
        data.append(row)
    
    return data

def extract_target_values(data):
    """Extract the specific values we want to copy to Excel"""
    distance_items = []
    int_circle_values = []
    
    for item in data:
        if item['Type'] == 'DISTANCE':
            if 'X' in item and 'ID' in item:
                distance_items.append({
                    'id': item['ID'], 
                    'value': round(abs(item['X']), 2)
                })
        elif item['Type'] == 'INT-CIRCLE':
            if 'Radius' in item:
                int_circle_values.append(round(item['Radius'], 2))
    
    distance_order = ['3', '2', '1', '4']
    distance_values = []
    
    for target_id in distance_order:
        for item in distance_items:
            if str(item['id']) == target_id:
                distance_values.append(item['value'])
                break
    
    return distance_values, int_circle_values

def extract_lot_prefix(lot_number):
    """Extract LOT prefix"""
    if not lot_number:
        return None
    match = re.match(r'^([A-Za-z]+\d+)', lot_number)
    return match.group(1) if match else None

def validate_date_format(date_string):
    """Validate date format YYYY/MM/DD"""
    if not date_string:
        return False
    return bool(re.match(r'^\d{4}/\d{2}/\d{2}$', date_string))

def update_excel_file(excel_file, distance_values, int_circle_values, distance_cells, int_circle_cells, 
                       lot_number=None, inspection_date=None, lot_prefix=None):
    """Update Excel file"""
    try:
        wb = load_workbook(excel_file)
        ws = wb["sheet"] if "sheet" in wb.sheetnames else wb.active
        st.info(f"✓ Using sheet: '{ws.title}'")
        
        def write_cell(cell_ref, value):
            try:
                cell = ws[cell_ref]
                if isinstance(cell, MergedCell):
                    for merged_range in ws.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            min_col, min_row, max_col, max_row = merged_range.bounds
                            ws.cell(row=min_row, column=min_col).value = value
                            return True
                else:
                    cell.value = value
                    return True
            except Exception as e:
                st.error(f"Error updating {cell_ref}: {e}")
                return False
        
        count = 0
        
        if lot_number and write_cell("B1", lot_number):
            count += 1
            st.success(f"✓ B1: {lot_number}")
        
        if inspection_date and write_cell("B2", inspection_date):
            count += 1
            st.success(f"✓ B2: {inspection_date}")
        
        if lot_prefix and write_cell("B3", lot_prefix):
            count += 1
            st.success(f"✓ B3: {lot_prefix}")
        
        for value, cell in zip(distance_values, distance_cells):
            if cell.strip() and write_cell(cell.strip(), value):
                count += 1
        
        for value, cell in zip(int_circle_values, int_circle_cells):
            if cell.strip() and write_cell(cell.strip(), value):
                count += 1
        
        st.info(f"📝 {count} cells updated")
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error: {e}")
        return None

def main():
    st.title("全検箇所測定データをExcelに転記するツール")
    
    if 'lot_number' not in st.session_state:
        st.session_state.lot_number = None
        st.session_state.inspection_date = None
        st.session_state.lot_prefix = None
    
    st.subheader("ステップ１：測定データファイルのアップロード")
    uploaded_file = st.file_uploader("ファイルを選択", type=["txt", "dat", "csv"])
    
    if uploaded_file:
        data = parse_data(uploaded_file.read())
        
        if data:
            st.success(f"✅ {len(data)}件のレコードを読み込みました")
            distance_values, int_circle_values = extract_target_values(data)
            st.info(f"📊 DISTANCE: {len(distance_values)}件, INT-CIRCLE: {len(int_circle_values)}件")
            
            if distance_values or int_circle_values:
                st.subheader("ステップ２：エクセルファイルの選択")
                excel_file = st.file_uploader("エクセルファイルを選択", type=["xlsx"])
                
                if excel_file:
                    st.subheader("ステップ３：セルの指定")
                    
                    option = st.radio("", ["デフォルト設定 (A列)", "カスタム指定"], index=0)
                    
                    if option == "デフォルト設定 (A列)":
                        distance_cells = [f"A{i+1}" for i in range(len(distance_values))]
                        int_circle_cells = [f"A{i+1+len(distance_values)}" for i in range(len(int_circle_values))]
                    else:
                        st.write("**DISTANCE cells:**")
                        distance_cells = [st.text_input(f"Value {distance_values[i]}", key=f"d{i}", placeholder=f"A{i+1}") 
                                        for i in range(len(distance_values))]
                        st.write("**INT-CIRCLE cells:**")
                        int_circle_cells = [st.text_input(f"Value {int_circle_values[i]}", key=f"c{i}", 
                                          placeholder=f"A{i+1+len(distance_values)}") 
                                          for i in range(len(int_circle_values))]
                    
                    st.subheader("ステップ４：LOT情報")
                    
                    with st.form("lot_form"):
                        lot_num = st.text_input("LOT番号", placeholder="例: LOT234(234-245)")
                        insp_date = st.text_input("検査日 (YYYY/MM/DD)", placeholder="例: 2025/10/07")
                        
                        if lot_num:
                            prefix = extract_lot_prefix(lot_num)
                            if prefix:
                                st.info(f"🔍 プレフィックス: **{prefix}**")
                        
                        st.write("**出力:** B1=LOT番号, B2=検査日, B3=プレフィックス")
                        
                        submitted = st.form_submit_button("✓ 確認")
                    
                    if submitted and lot_num and insp_date:
                        if validate_date_format(insp_date):
                            st.session_state.lot_number = lot_num
                            st.session_state.inspection_date = insp_date
                            st.session_state.lot_prefix = extract_lot_prefix(lot_num)
                            st.success("✅ LOT情報を保存しました")
                        else:
                            st.error("❌ 日付形式: YYYY/MM/DD")
                    
                    if st.button("📊 エクセル更新", type="primary"):
                        if not st.session_state.lot_number:
                            st.error("❌ LOT情報を入力してください")
                        else:
                            with st.spinner("更新中..."):
                                result = update_excel_file(
                                    excel_file, distance_values, int_circle_values,
                                    distance_cells, int_circle_cells,
                                    st.session_state.lot_number,
                                    st.session_state.inspection_date,
                                    st.session_state.lot_prefix
                                )
                            
                            if result:
                                st.success("✅ 完了!")
                                filename = f"水平ノズル{st.session_state.lot_number}全箇所測定{st.session_state.inspection_date}.xlsx"
                                st.download_button(
                                    "💾 ダウンロード",
                                    result.getvalue(),
                                    filename,
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
        else:
            st.error("データが見つかりません")
    else:
        st.info("👆 ファイルをアップロードしてください")

if __name__ == "__main__":
    main()
