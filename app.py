\import streamlit as st
import pandas as pd
import numpy as np
import re
import base64
from io import BytesIO
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import tempfile
import os

st.set_page_config(page_title="Geometric Data Viewer", layout="wide")

def parse_data(file_path):
    with open(file_path, 'r') as file:
        lines = file.readlines()
    
    # Create lists to store data
    data = []
    
    for line in lines:
        # Split the line by semicolons
        parts = line.strip().split(';')
        
        if len(parts) < 2:
            continue
        
        # Extract ID and object type
        obj_id = parts[0]
        obj_type = parts[1]
        
        # Initialize a dictionary for the row
        row = {'ID': obj_id, 'Type': obj_type}
        
        # Extract remaining values
        remaining_values = [p.strip() for p in parts[2:] if p.strip()]
        
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
        else:
            # Generic handling for unknown types
            cols = [f'Val{i+1}' for i in range(len(remaining_values))]
        
        # Add values to the row dictionary
        for i, val in enumerate(remaining_values):
            if i < len(cols) and cols[i]:
                try:
                    # Try to convert to float if possible
                    row[cols[i]] = float(val) if re.match(r'^-?\d+\.?\d*$', val) else val
                except:
                    row[cols[i]] = val
        
        data.append(row)
    
    return data

def create_pdf(data, filename="geometric_data.pdf"):
    """
    Create a PDF file from the geometric data
    """
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter))
    elements = []
    
    # Add title
    styles = getSampleStyleSheet()
    elements.append(Paragraph("Geometric Data Report", styles['Title']))
    elements.append(Spacer(1, 12))
    
    # Group data by object type
    object_types = set(item['Type'] for item in data)
    
    for obj_type in sorted(object_types):
        # Add section title for each object type
        elements.append(Paragraph(f"{obj_type} Objects", styles['Heading2']))
        elements.append(Spacer(1, 6))
        
        # Filter data for current object type
        type_data = [item for item in data if item['Type'] == obj_type]
        
        if not type_data:
            continue
            
        # Get all keys from all items
        all_keys = set()
        for item in type_data:
            all_keys.update(item.keys())
        
        # Sort keys to ensure "ID" and "Type" come first
        sorted_keys = ["ID", "Type"] + sorted([k for k in all_keys if k not in ["ID", "Type"]])
        
        # Create table data
        table_data = [sorted_keys]  # Header row
        
        for item in type_data:
            row = []
            for key in sorted_keys:
                row.append(str(item.get(key, "")))
            table_data.append(row)
        
        # Create table
        table = Table(table_data)
        
        # Style the table
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])
        
        # Apply alternating row colors
        for i in range(1, len(table_data)):
            if i % 2 == 0:
                style.add('BACKGROUND', (0, i), (-1, i), colors.lightgrey)
                
        table.setStyle(style)
        elements.append(table)
        elements.append(Spacer(1, 12))
    
    # Build PDF
    doc.build(elements)
    
    # Get PDF data
    pdf_data = buffer.getvalue()
    buffer.close()
    
    return pdf_data

def get_download_link(pdf_data, filename="geometric_data.pdf"):
    """
    Generate a download link for the PDF file
    """
    b64 = base64.b64encode(pdf_data).decode()
    href = f'<a href="data:application/pdf;base64,{b64}" download="{filename}">Download PDF Report</a>'
    return href

def main():
    st.title("Geometric Data Visualization")
    
    st.write("This application displays the geometric data from an input file.")
    
    # File upload option with prominent UI
    st.subheader("Data Input")
    uploaded_file = st.file_uploader("Upload your geometric data file (.txt)", type="txt")
    
    use_sample_data = st.checkbox("Use sample data (341.txt)", value=(uploaded_file is None))
    
    if uploaded_file is not None:
        # Use uploaded file
        with open("temp_data.txt", "wb") as f:
            f.write(uploaded_file.getbuffer())
        file_path = "temp_data.txt"
        st.success(f"Successfully loaded: {uploaded_file.name}")
    elif use_sample_data:
        # Use default file path
        file_path = "341.txt"
        st.info("Using sample data from 341.txt")
    else:
        st.warning("Please upload a file or check the 'Use sample data' option")
        return
    
    # Parse the data
    data = parse_data(file_path)
    
    # Group by object type
    object_types = set(item['Type'] for item in data)
    
    # Create a multi-tab display
    st.write(f"**Found {len(object_types)} different object types in the data**")
    
    # Convert to DataFrame
    df = pd.DataFrame(data)
    
    # Show all data in a single table
    st.subheader("All Geometric Data")
    st.dataframe(df, use_container_width=True)
    
    # Show data by type in separate tabs
    st.subheader("Data by Object Type")
    tabs = st.tabs([obj_type for obj_type in sorted(object_types)])
    
    for i, obj_type in enumerate(sorted(object_types)):
        with tabs[i]:
            type_data = [item for item in data if item['Type'] == obj_type]
            type_df = pd.DataFrame(type_data)
            st.dataframe(type_df, use_container_width=True)
            
            # Add basic statistics for numeric columns
            numeric_cols = type_df.select_dtypes(include=[np.number]).columns.tolist()
            if numeric_cols:
                st.subheader(f"Statistics for {obj_type}")
                st.dataframe(type_df[numeric_cols].describe(), use_container_width=True)
        
        # Export to PDF section
        st.subheader("Export Data")
        
        # Create PDF from the data
        pdf_data = create_pdf(data)
        
        # Provide download link
        st.markdown(get_download_link(pdf_data), unsafe_allow_html=True)
        
        # Add export options
        col1, col2, col3 = st.columns(3)
        
        # Option to also export as CSV
        csv = df.to_csv(index=False)
        b64_csv = base64.b64encode(csv.encode()).decode()
        with col1:
            st.markdown(
                f'<a href="data:file/csv;base64,{b64_csv}" download="geometric_data.csv">Download CSV</a>',
                unsafe_allow_html=True
            )
        
        # Option to export as Excel
        with col2:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Write each object type to a different sheet
                for obj_type in sorted(object_types):
                    type_data = [item for item in data if item['Type'] == obj_type]
                    type_df = pd.DataFrame(type_data)
                    type_df.to_excel(writer, sheet_name=obj_type[:31], index=False)  # Excel sheet names limited to 31 chars
                
                # Add a sheet with all data
                df.to_excel(writer, sheet_name='All Data', index=False)
            
            b64_excel = base64.b64encode(buffer.getvalue()).decode()
            st.markdown(
                f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="geometric_data.xlsx">Download Excel</a>',
                unsafe_allow_html=True
            )
        
        # Add some information about the export
        with col3:
            st.info("PDF includes data tables for each object type")

if __name__ == "__main__":
    main()