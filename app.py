import streamlit as st
import pandas as pd
import openpyxl
import io

# 1. Page Configuration
st.set_page_config(page_title="Overview of Office Spaces", layout="wide")

st.title("🏢 Overview of Office Spaces & Areas")

# 2. Upload File
uploaded_file = st.file_uploader("Upload your Excel file (e.g. Test 1.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # Shift index to match Excel row numbering (+1)
    df.index = df.index + 1

    # Ensure numeric columns are true floats
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)

    # -------------------------------------------------------------
    # CALCULATE TOTALS ROW FOR STREAMLIT DISPLAY
    # -------------------------------------------------------------
    df_display = df.copy()
    
    # Create sum dictionary for all numeric columns
    numeric_cols = df_display.select_dtypes(include=['float', 'int']).columns
    sums = df_display[numeric_cols].sum()
    
    # Create Total Row
    total_row = pd.DataFrame(sums).T
    total_row[df_display.columns[0]] = "Celkem"
    total_row.index = ["Total"]
    
    # Append Total row to the bottom
    df_with_total = pd.concat([df_display, total_row])
    
    # FIX: Cast index to string so PyArrow handles integer rows + "Total" cleanly
    df_with_total.index = df_with_total.index.astype(str)

    st.subheader("Uploaded Data with Totals")

    # Helper function: Format numbers with Czech separators, hide zeros (render as blank)
    def format_czech_hide_zeros(val):
        if pd.isna(val) or not isinstance(val, (int, float)):
            return val
        if val == 0:
            return ""
        return f"{val:,.2f}".replace(",", " ").replace(".", ",")

    # Display table in Streamlit
    st.dataframe(
        df_with_total.style.format(format_czech_hide_zeros),
        width="stretch",
        height='content'
    )

    # -------------------------------------------------------------
    # EXCEL EXPORT WITH DYNAMIC EXCEL =SUM() FORMULAS
    # -------------------------------------------------------------
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Office Overview', index=True)
        
        ws = writer.sheets['Office Overview']
        
        # Header Styling
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = openpyxl.styles.Font(bold=True)

        # Format Data Rows
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if col == 1:
                    cell.alignment = openpyxl.styles.Alignment(horizontal='center')
                elif isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00;-#,##0.00;""'
                    cell.alignment = openpyxl.styles.Alignment(horizontal='right')

        # Add Total Row in Excel at bottom
        total_excel_row = ws.max_row + 1
        ws.cell(row=total_excel_row, column=2, value="Celkem (Total)").font = openpyxl.styles.Font(bold=True)
        
        # Insert =SUM(C2:C19) formulas across all numeric columns
        for col in range(3, ws.max_column + 1):
            col_letter = openpyxl.utils.get_column_letter(col)
            formula = f"=SUM({col_letter}2:{col_letter}{total_excel_row - 1})"
            
            sum_cell = ws.cell(row=total_excel_row, column=col, value=formula)
            sum_cell.font = openpyxl.styles.Font(bold=True)
            sum_cell.number_format = '#,##0.00;-#,##0.00;""'
            sum_cell.alignment = openpyxl.styles.Alignment(horizontal='right')
            
            # Add top thin line & bottom double border for Excel accounting style
            sum_cell.border = openpyxl.styles.Border(
                top=openpyxl.styles.Side(style='thin'),
                bottom=openpyxl.styles.Side(style='double')
            )

        # Auto-adjust column widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    buffer.seek(0)

    st.download_button(
        label="📥 Download Overview with Totals (.xlsx)",
        data=buffer,
        file_name="Office_Overview_With_Totals.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload your Excel file to display the table.")
