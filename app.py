import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
import openpyxl.cell.cell
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
                try:
                    numeric_pattern = r'^-?\d+\.?\d*$'
                    if re.match(numeric_pattern, val):
                        row[cols[i]] = float(val)
                    else:
                        row[cols[i]] = val
                except:
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
    """Extract LOT prefix from LOT number string"""
    if not lot_number:
        return None
    
    match = re.match(r'^([A-Za-z]+\d+)', lot_number)
    if match:
        return match.group(1)
    return None

def validate_date_format(date_string):
    """Validate date format YYYY/MM/DD"""
    if not date_string:
        return False
    
    pattern = r'^\d{4}/\d{2}/\d{2}$'
    return bool(re.match(pattern, date_string))

def update_excel_file(excel_file, distance_values, int_circle_values, distance_cells, int_circle_cells, 
                       lot_number=None, inspection_date=None, lot_prefix=None):
    """Update Excel file with the extracted values"""
    try:
        wb = load_workbook(excel_file)
        
        # Use "sheet" if it exists, otherwise use active sheet
        if "sheet" in wb.sheetnames:
            ws = wb["sheet"]
            st.info(f"✓ Using sheet: 'sheet'")
        else:
            ws = wb.active
            st.info(f"✓ Using active sheet: '{ws.title}'")
        
        def write_to_cell(worksheet, cell_ref, value):
            """Write to a cell, handling merged cells"""
            try:
                cell = worksheet[cell_ref]
                
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    for merged_range in worksheet.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            min_col, min_row, max_col, max_row = merged_range.bounds
                            top_left_cell = worksheet.cell(row=min_row, column=min_col)
                            top_left_cell.value = value
                            return True
                else:
                    cell.value = value
                    return True
            except Exception as e:
                st.error(f"❌ Error updating cell {cell_ref}: {e}")
                return False
        
        successful_writes = 0
        
        # Write LOT information to B column
        if lot_number:
            if write_to_cell(ws, "B1", lot_number):
                successful_writes += 1
                st.success(f"✓ Wrote LOT番号 '{lot_number}' to cell B1")
        
        if inspection_date:
            if write_to_cell(ws, "B2", inspection_date):
                successful_writes += 1
                st.success(f"✓ Wrote 検査日 '{inspection_date}' to cell B2")
        
        if lot_prefix:
            if write_to_cell(ws, "B3", lot_prefix):
                successful_writes += 1
                st.success(f"✓ Wrote LOTプレフィックス '{lot_prefix}' to cell B3")
        
        # Update DISTANCE values
        for i, (value, cell) in enumerate(zip(distance_values, distance_cells)):
            if cell.strip():
                if write_to_cell(ws, cell.strip(), value):
                    successful_writes += 1
        
        # Update INT-CIRCLE values
        for i, (value, cell) in enumerate(zip(int_circle_values, int_circle_cells)):
            if cell.strip():
                if write_to_cell(ws, cell.strip(), value):
                    successful_writes += 1
        
        st.info(f"📝 Total cells updated: {successful_writes}")
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output
    
    except Exception as e:
        st.error(f"Error processing Excel file: {e}")
        return None

def main():
    st.title("全検箇所測定データをExcelに転記するツール")
    
    # Initialize session state
    if 'lot_number' not in st.session_state:
        st.session_state['lot_number'] = None
        st.session_state['inspection_date'] = None
        st.session_state['lot_prefix'] = None
    
    st.subheader("ステップ１：測定データファイルのアップロード")
    st.write("以下の形式のファイルをアップロードしてください。")
    uploaded_file = st.file_uploader(
        "ファイルを選択", 
        type=["txt", "dat", "csv"],
        help="セミコロン区切りのデータファイルを選択してください"
    )
    
    if uploaded_file is not None:
        file_content = uploaded_file.read()
        data = parse_data(file_content)
        
        if data:
            st.success(f"✅ {len(data)}件のレコードを読み込みました")
            
            distance_values, int_circle_values = extract_target_values(data)
            
            st.info(f"📊 抽出されたデータ: DISTANCE値 {len(distance_values)}件, INT-CIRCLE値 {len(int_circle_values)}件")
            
            if distance_values or int_circle_values:
                st.subheader("ステップ２：アップロードしたエクセルファイルの指定セルにデータを出力します。")
                
                excel_file = st.file_uploader(
                    "エクセルファイルを選択",
                    type=["xlsx", "xls"],
                    help="更新したいExcelファイルを選択してください"
                )
                
                if excel_file is not None:
                    st.subheader("ステップ３：セルの指定（オプション）")
                    
                    location_option = st.radio(
                        "データを移行するセルの指定:",
                        ["デフォルト設定 (A列に出力)", "カスタム指定"],
                        index=0
                    )
                    
                    if location_option == "デフォルト設定 (A列に出力)":
                        distance_cells = [f"A{i+1}" for i in range(len(distance_values))]
                        int_circle_cells = [f"A{i+1+len(distance_values)}" for i in range(len(int_circle_values))]
                        has_distance_cells = True
                        has_int_circle_cells = True
                    else:
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write("**DISTANCE Values Cell References:**")
                            distance_cells = []
                            for i in range(len(distance_values)):
                                cell = st.text_input(
                                    f"Cell for DISTANCE value {i+1} ({distance_values[i]})",
                                    key=f"distance_cell_{i}",
                                    placeholder=f"e.g., A{i+1}"
                                )
                                distance_cells.append(cell)
                        
                        with col2:
                            st.write("**INT-CIRCLE Values Cell References:**")
                            int_circle_cells = []
                            for i in range(len(int_circle_values)):
                                cell = st.text_input(
                                    f"Cell for INT-CIRCLE value {i+1} ({int_circle_values[i]})",
                                    key=f"int_circle_cell_{i}",
                                    placeholder=f"e.g., A{i+1+len(distance_values)}"
                                )
                                int_circle_cells.append(cell)
                        
                        has_distance_cells = any(cell.strip() for cell in distance_cells)
                        has_int_circle_cells = any(cell.strip() for cell in int_circle_cells)
                    
                    # LOT Information Input Section
                    st.subheader("ステップ４：LOT情報の入力")
                    st.write("エクセルファイルのB列に出力されるLOT情報を入力してください。")
                    
                    with st.form(key="lot_info_form"):
                        lot_number = st.text_input(
                            "LOT番号を入力してください",
                            placeholder="例: LOT234(234-245)",
                            help="LOT番号を入力してください（例: LOT234(234-245)）"
                        )
                        
                        inspection_date = st.text_input(
                            "検査日を入力してください (YYYY/MM/DD)",
                            placeholder="例: 2025/10/07",
                            help="検査日を YYYY/MM/DD 形式で入力してください"
                        )
                        
                        if lot_number:
                            lot_prefix = extract_lot_prefix(lot_number)
                            if lot_prefix:
                                st.info(f"🔍 自動抽出されたLOTプレフィックス: **{lot_prefix}** (B3に出力されます)")
                            else:
                                st.warning("⚠️ LOTプレフィックスを抽出できませんでした")
                        
                        st.write("**出力先:**")
                        st.write("• LOT番号 → **B1**")
                        st.write("• 検査日 → **B2**")
                        st.write("• LOTプレフィックス（自動抽出） → **B3**")
                        
                        submit_button = st.form_submit_button("✓ 確認", type="secondary")
                    
                    if submit_button:
                        if lot_number and inspection_date:
                            if not validate_date_format(inspection_date):
                                st.error("❌ 検査日は YYYY/MM/DD 形式で入力してください（例: 2025/10/07）")
                            else:
                                lot_prefix = extract_lot_prefix(lot_number)
                                if not lot_prefix:
                                    st.warning("⚠️ LOTプレフィックスを抽出できませんでした。LOT番号を確認してください。")
                                
                                st.success("✅ LOT情報が入力されました！下のボタンをクリックしてエクセルファイルを更新してください。")
                                
                                st.session_state['lot_number'] = lot_number
                                st.session_state['inspection_date'] = inspection_date
                                st.session_state['lot_prefix'] = lot_prefix
                        else:
                            st.error("❌ LOT番号と検査日の両方を入力してください。")
                    
                    # Update Excel button
                    if st.button("📊 エクセルファイルの更新", type="primary"):
                        if not st.session_state.get('lot_number') or not st.session_state.get('inspection_date'):
                            st.error("❌ LOT情報を入力してから更新ボタンをクリックしてください。")
                        elif (location_option == "デフォルト設定 (A列に出力)") or (has_distance_cells or has_int_circle_cells):
                            with st.spinner("エクセルファイルを更新中..."):
                                updated_excel = update_excel_file(
                                    excel_file, 
                                    distance_values, 
                                    int_circle_values, 
                                    distance_cells, 
                                    int_circle_cells,
                                    st.session_state['lot_number'],
                                    st.session_state['inspection_date'],
                                    st.session_state['lot_prefix']
                                )
                            
                            if updated_excel:
                                st.success("✅ エクセルファイルの更新が完了しました!")
                                
                                lot_num = st.session_state['lot_number']
                                insp_date = st.session_state['inspection_date']
                                dynamic_filename = f"水平ノズル{lot_num}全箇所測定{insp_date}.xlsx"
                                
                                st.download_button(
                                    label="💾 更新したエクセルファイルのダウンロード",
                                    data=updated_excel.getvalue(),
                                    file_name=dynamic_filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        else:
                            st.warning("⚠️ カスタム指定の場合は、少なくとも1つのセル参照を入力してください。")
        else:
            st.error("有効なデータが見つかりませんでした")
    else:
        st.info("👆 Browse File ボタンを押して処理するテキストデータをアップロードしてください")

if __name__ == "__main__":
    main()