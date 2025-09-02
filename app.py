import streamlit as st
import pandas as pd
import re
from openpyxl import load_workbook
from io import BytesIO
import tempfile
import os

st.set_page_config(page_title="Geometric Data Reader", layout="wide")

def parse_data(file_content):
    """Parse data from file content string"""
    if isinstance(file_content, bytes):
        lines = file_content.decode('utf-8').splitlines()
    else:
        lines = file_content.splitlines()
    
    # Create list to store data
    data = []
    
    for line in lines:
        # Skip empty lines
        if not line.strip():
            continue
            
        # Split the line by semicolons
        parts = line.strip().split(';')
        
        if len(parts) < 2:
            continue
        
        # Extract ID and object type
        obj_id = parts[0]
        obj_type = parts[1]
        
        # Initialize a dictionary for the row
        row = {'ID': obj_id, 'Type': obj_type}
        
        # Extract remaining values (keep empty strings to maintain column mapping)
        remaining_values = [p.strip() for p in parts[2:]]
        
        # Determine column names based on object type
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
            # Generic handling for unknown types
            cols = [f'Val{i+1}' for i in range(len(remaining_values))]
        
        # Add values to the row dictionary
        for i, val in enumerate(remaining_values):
            if i < len(cols) and cols[i] and val:
                try:
                    # Try to convert to float if it's a number
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
    
    # Collect all DISTANCE and INT-CIRCLE items
    for item in data:
        if item['Type'] == 'DISTANCE':
            if 'X' in item and 'ID' in item:
                distance_items.append({
                    'id': item['ID'], 
                    'value': round(abs(item['X']), 2)
                })
        elif item['Type'] == 'INT-CIRCLE':
            # For INT-CIRCLE, we want the Radius column value (not K)
            if 'Radius' in item:
                int_circle_values.append(round(item['Radius'], 2))
    
    # Sort DISTANCE values in the specific order: ID 3, 2, 1, 4
    distance_order = ['3', '2', '1', '4']
    distance_values = []
    
    for target_id in distance_order:
        for item in distance_items:
            if str(item['id']) == target_id:
                distance_values.append(item['value'])
                break
    
    return distance_values, int_circle_values

def update_excel_file(excel_file, distance_values, int_circle_values, distance_cells, int_circle_cells):
    """Update Excel file with the extracted values"""
    try:
        # Load the workbook
        wb = load_workbook(excel_file)
        
        # Get the active sheet (or you can specify sheet name)
        ws = wb.active
        
        # Update DISTANCE values
        for i, (value, cell) in enumerate(zip(distance_values, distance_cells)):
            if cell.strip():  # Only update if cell reference is provided
                try:
                    ws[cell.strip()] = value
                except Exception as e:
                    st.error(f"Error updating cell {cell}: {e}")
        
        # Update INT-CIRCLE values
        for i, (value, cell) in enumerate(zip(int_circle_values, int_circle_cells)):
            if cell.strip():  # Only update if cell reference is provided
                try:
                    ws[cell.strip()] = value
                except Exception as e:
                    st.error(f"Error updating cell {cell}: {e}")
        
        # Save to BytesIO object
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output
    
    except Exception as e:
        st.error(f"Error processing Excel file: {e}")
        return None

def main():
    st.title("Geometric Data Reader with Excel Export")
    st.write("Browse and select a text file with geometric data, then export specific values to Excel.")
    
    # File upload with more detailed instructions
    st.subheader("Step 1: Select Geometric Data File")
    st.write("Click 'Browse files' below to select your geometric data file:")
    uploaded_file = st.file_uploader(
        "Choose your geometric data file", 
        type=["txt", "dat", "csv"],
        help="Select a .txt, .dat, or .csv file containing geometric data from your local folders"
    )
    
    if uploaded_file is not None:
        # Read and parse the file
        file_content = uploaded_file.read()
        data = parse_data(file_content)
        
        if data:
            st.success(f"Successfully loaded {len(data)} records from {uploaded_file.name}")
            
            # Convert to DataFrame
            df = pd.DataFrame(data)
            
            # Show all data
            st.subheader("All Data")
            st.dataframe(df, use_container_width=True)
            
            # Extract target values
            distance_values, int_circle_values = extract_target_values(data)
            
            # Show extracted values
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📏 DISTANCE Values")
                if distance_values:
                    for i, val in enumerate(distance_values, 1):
                        st.write(f"Value {i}: **{val}**")
                else:
                    st.write("No DISTANCE values found")
            
            with col2:
                st.subheader("🔵 INT-CIRCLE Values")
                if int_circle_values:
                    for i, val in enumerate(int_circle_values, 1):
                        st.write(f"Value {i}: **{val}**")
                else:
                    st.write("No INT-CIRCLE values found")
            
            # Excel export section
            if distance_values or int_circle_values:
                st.subheader("Step 2: Export to Excel")
                
                # Upload Excel file
                excel_file = st.file_uploader(
                    "Upload Excel file to update",
                    type=["xlsx", "xls"],
                    help="Select the Excel file where you want to paste the values"
                )
                
                if excel_file is not None:
                    st.subheader("Step 3: Cell Location Settings")
                    
                    # Option to choose between automatic BE column or custom cells
                    location_option = st.radio(
                        "Choose how to place the values:",
                        ["Automatic (Column BE, rows 1-6)", "Custom cell references"],
                        index=0
                    )
                    
                    if location_option == "Automatic (Column BE, rows 1-6)":
                        # Automatically set cells to BE1-BE6
                        distance_cells = [f"BE{i+1}" for i in range(len(distance_values))]
                        int_circle_cells = [f"BE{i+1+len(distance_values)}" for i in range(len(int_circle_values))]
                        
                        st.write("**Values will be placed in:**")
                        for i, val in enumerate(distance_values):
                            st.write(f"• DISTANCE value {val} → **BE{i+1}**")
                        for i, val in enumerate(int_circle_values):
                            st.write(f"• INT-CIRCLE value {val} → **BE{i+1+len(distance_values)}**")
                        
                        has_distance_cells = True
                        has_int_circle_cells = True
                        
                    else:
                        # Custom cell input (original functionality)
                        st.write("Enter the cell references where you want to paste the values (e.g., A1, B2, C3):")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write("**DISTANCE Values Cell References:**")
                            distance_cells = []
                            for i in range(len(distance_values)):
                                cell = st.text_input(
                                    f"Cell for DISTANCE value {i+1} ({distance_values[i]})",
                                    key=f"distance_cell_{i}",
                                    placeholder=f"e.g., BE{i+1}"
                                )
                                distance_cells.append(cell)
                        
                        with col2:
                            st.write("**INT-CIRCLE Values Cell References:**")
                            int_circle_cells = []
                            for i in range(len(int_circle_values)):
                                cell = st.text_input(
                                    f"Cell for INT-CIRCLE value {i+1} ({int_circle_values[i]})",
                                    key=f"int_circle_cell_{i}",
                                    placeholder=f"e.g., BE{i+1+len(distance_values)}"
                                )
                                int_circle_cells.append(cell)
                        
                        # Check if at least one cell reference is provided
                        has_distance_cells = any(cell.strip() for cell in distance_cells)
                        has_int_circle_cells = any(cell.strip() for cell in int_circle_cells)
                    
                    # Update Excel button
                    if st.button("📊 Update Excel File", type="primary"):
                        if (location_option == "Automatic (Column BE, rows 1-6)") or (has_distance_cells or has_int_circle_cells):
                            updated_excel = update_excel_file(
                                excel_file, 
                                distance_values, 
                                int_circle_values, 
                                distance_cells, 
                                int_circle_cells
                            )
                            
                            if updated_excel:
                                st.success("✅ Excel file updated successfully!")
                                
                                # Download button
                                st.download_button(
                                    label="💾 Download Updated Excel File",
                                    data=updated_excel.getvalue(),
                                    file_name=f"updated_{excel_file.name}",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                                
                                # Show summary
                                st.subheader("Update Summary")
                                for i, (val, cell) in enumerate(zip(distance_values, distance_cells)):
                                    if cell.strip():
                                        st.write(f"✓ DISTANCE value {val} → Cell {cell}")
                                
                                for i, (val, cell) in enumerate(zip(int_circle_values, int_circle_cells)):
                                    if cell.strip():
                                        st.write(f"✓ INT-CIRCLE value {val} → Cell {cell}")
                        else:
                            st.warning("⚠️ Please provide at least one cell reference in custom mode.")
            
            # Group by object type and show in separate sections
            object_types = set(item['Type'] for item in data)
            
            if len(object_types) > 1:
                st.subheader("Data by Object Type")
                
                for obj_type in sorted(object_types):
                    st.write(f"**{obj_type}**")
                    type_data = [item for item in data if item['Type'] == obj_type]
                    type_df = pd.DataFrame(type_data)
                    st.dataframe(type_df, use_container_width=True)
                    st.write("---")
        else:
            st.error("No valid data found in the uploaded file")
    else:
        st.info("👆 Click 'Browse files' above to select a file from your computer")
        
        # Show file format example
        st.subheader("Expected File Format")
        st.write("Your file should contain semicolon-separated data like this:")
        st.code("""ID1;PLANE;Method;X;Y;Z;A;B;C;;D;Dev
ID2;CIRCLE;Method;X;Y;Z;I;J;K;;Radius;Dev
ID3;PT-COMP;Method;X;Y;Z
ID4;DISTANCE;;X;Y;Z;;;;;Distance
ID5;INT-CIRCLE;;X;Y;Z;I;J;K;;Radius""", language="text")
        
        st.write("**Supported object types:** PLANE, CIRCLE, PT-COMP, DISTANCE, CONE, INT-CIRCLE, SYM-POINT")

if __name__ == "__main__":
    main()