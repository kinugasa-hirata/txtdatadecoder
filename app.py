import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import os

st.set_page_config(page_title="Geometric Data Reader", layout="wide")

# Path to the template Excel file in the project
TEMPLATE_EXCEL_PATH = "LOT追加測定箇所.xlsx"

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
        # Load workbook - handle both file path and uploaded file
        if isinstance(excel_file, str):
            wb = load_workbook(excel_file)
        else:
            wb = load_workbook(excel_file)
        
        ws = wb["sheet"] if "sheet" in wb.sheetnames else wb.active
        
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
            except:
                return False
        
        count = 0
        
        if lot_number:
            if write_cell("B1", lot_number):
                count += 1
        
        if inspection_date:
            if write_cell("B2", inspection_date):
                count += 1
        
        if lot_prefix:
            if write_cell("B3", lot_prefix):
                count += 1
        
        for value, cell in zip(distance_values, distance_cells):
            if cell.strip() and write_cell(cell.strip(), value):
                count += 1
        
        for value, cell in zip(int_circle_values, int_circle_cells):
            if cell.strip() and write_cell(cell.strip(), value):
                count += 1
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"エラー: {e}")
        return None

def main():
    st.title("全検箇所測定データをExcelに転記するツール")
    
    # Show info about template file
    with st.expander("ℹ️ セットアップ情報", expanded=False):
        st.write("""
        **テンプレートファイルの設定:**
        
        プロジェクトフォルダに `LOT追加測定箇所.xlsx` を配置すると、自動的に使用されます。
        
        テンプレートファイルがない場合は、手動でアップロードしてください。
        """)
    
    st.subheader("ステップ１：測定データファイルのアップロード")
    uploaded_file = st.file_uploader("ファイルを選択", type=["txt", "dat", "csv"])
    
    if uploaded_file:
        data = parse_data(uploaded_file.read())
        
        if data:
            st.success(f"✅ {len(data)}件のレコードを読み込みました")
            distance_values, int_circle_values = extract_target_values(data)
            st.info(f"📊 DISTANCE: {len(distance_values)}件, INT-CIRCLE: {len(int_circle_values)}件")
            
            if distance_values or int_circle_values:
                # Check if template Excel file exists
                excel_file = None
                use_template = False
                
                if os.path.exists(TEMPLATE_EXCEL_PATH):
                    st.success(f"✅ テンプレートファイルを使用します")
                    use_template = True
                    excel_file = TEMPLATE_EXCEL_PATH
                else:
                    st.info(f"📁 テンプレートファイルが見つかりません。Excelファイルをアップロードしてください。")
                    st.subheader("ステップ２：エクセルファイルの選択")
                    uploaded_excel = st.file_uploader("エクセルファイルを選択", type=["xlsx"])
                    if uploaded_excel:
                        excel_file = uploaded_excel
                
                if excel_file:
                    step_num = "ステップ２" if not use_template else "ステップ２"
                    st.subheader(f"{step_num}：セルの指定")
                    
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
                    
                    step_num = "ステップ３" if not use_template else "ステップ３"
                    st.subheader(f"{step_num}：LOT情報")
                    
                    lot_num = st.text_input("LOT番号を入力", placeholder="例: LOT234(234-245)")
                    insp_date = st.text_input("検査日を入力 (YYYY/MM/DD)", placeholder="例: 2025/10/07")
                    
                    st.write("**出力先:** B1=LOT番号, B2=検査日, B3=プレフィックス（自動抽出）")
                    
                    if st.button("📊 エクセル更新", type="primary"):
                        if not lot_num or not insp_date:
                            st.error("❌ LOT番号と検査日を入力してください")
                        elif not validate_date_format(insp_date):
                            st.error("❌ 検査日の形式が正しくありません (YYYY/MM/DD)")
                        else:
                            # Compute prefix automatically
                            lot_prefix = extract_lot_prefix(lot_num)
                            
                            with st.spinner("エクセルファイルを更新中..."):
                                result = update_excel_file(
                                    excel_file, distance_values, int_circle_values,
                                    distance_cells, int_circle_cells,
                                    lot_num, insp_date, lot_prefix
                                )
                            
                            if result:
                                st.success("✅ 更新完了!")
                                filename = f"水平ノズル{lot_num}全箇所測定{insp_date}.xlsx"
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
