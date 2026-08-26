import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Attendance Calculator", page_icon="📊", layout="wide")

st.title("📊 Attendance Calculator")
st.markdown("Upload your raw Excel attendance report to automatically calculate and view summary statistics")

# File upload
uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # Read Excel file using openpyxl engine
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None, engine='openpyxl')
        
        st.success("✅ File uploaded successfully!")
        
        # Parse attendance data
        employees = {}
        overall_stats = {
            'totalDays': 0,
            'presentDays': 0,
            'absentDays': 0,
            'leaves': 0,
            'weeklyOff': 0
        }
        
        current_employee = None
        is_data_section = False
        
        # Iterate through rows
        for i in range(len(df)):
            row = df.iloc[i]
            
            # Get first column value
            first_col = str(row.iloc[0] if len(row) > 0 else '').strip().lower()
            
            # Look for employee code marker
            if 'emp code' in first_col:
                if i + 1 < len(df):
                    next_row = df.iloc[i + 1]
                    if len(next_row) > 6:
                        emp_code = str(next_row.iloc[4]).strip() if pd.notna(next_row.iloc[4]) else "Unknown"
                        emp_name = str(next_row.iloc[6]).strip() if pd.notna(next_row.iloc[6]) else f"Employee {emp_code}"
                        
                        current_employee = {
                            'code': emp_code,
                            'name': emp_name,
                            'stats': {
                                'present': 0,
                                'absent': 0,
                                'leave': 0,
                                'weeklyOff': 0,
                                'total': 0
                            }
                        }
                        employees[emp_code] = current_employee
            
            # Look for attendance date marker
            if 'att. date' in first_col or 'att.date' in first_col:
                is_data_section = True
                continue
            
            # Process attendance records
            if is_data_section and current_employee and len(row) > 12:
                if pd.notna(row.iloc[12]):
                    status = str(row.iloc[12]).strip()
                    current_employee['stats']['total'] += 1
                    
                    if status == 'Present':
                        current_employee['stats']['present'] += 1
                        overall_stats['presentDays'] += 1
                    elif status == 'Absent':
                        current_employee['stats']['absent'] += 1
                        overall_stats['absentDays'] += 1
                    elif status == 'WeeklyOff':
                        current_employee['stats']['weeklyOff'] += 1
                        overall_stats['weeklyOff'] += 1
                    elif 'Leave' in status:
                        current_employee['stats']['leave'] += 1
                        overall_stats['leaves'] += 1
            
            # End data section
            if 'total duration' in first_col:
                is_data_section = False
        
        # Calculate totals
        overall_stats['totalDays'] = (overall_stats['presentDays'] + 
                                      overall_stats['absentDays'] + 
                                      overall_stats['leaves'] + 
                                      overall_stats['weeklyOff'])
        
        # Display Overall Summary
        st.subheader("📈 Overall Summary")
        
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        
        with col1:
            st.metric("Working Days", overall_stats['totalDays'])
        with col2:
            st.metric("Present", overall_stats['presentDays'])
        with col3:
            st.metric("Absent", overall_stats['absentDays'])
        with col4:
            st.metric("Leaves", overall_stats['leaves'])
        with col5:
            st.metric("Weekly Off", overall_stats['weeklyOff'])
        with col6:
            if overall_stats['totalDays'] > 0:
                attendance_pct = round((overall_stats['presentDays'] / overall_stats['totalDays'] * 100))
            else:
                attendance_pct = 0
            st.metric("Attendance %", f"{attendance_pct}%")
        
        # Display Employee Details
        st.subheader("👥 Employee Details")
        
        if len(employees) > 0:
            # Sort employees by name
            sorted_employees = sorted(employees.values(), key=lambda x: x['name'])
            
            # Create table data
            table_data = []
            for emp in sorted_employees:
                if emp['stats']['total'] > 0:
                    attendance = round((emp['stats']['present'] / emp['stats']['total'] * 100))
                else:
                    attendance = 0
                    
                table_data.append({
                    'Employee Name': emp['name'],
                    'Employee Code': emp['code'],
                    'Present': emp['stats']['present'],
                    'Absent': emp['stats']['absent'],
                    'Leave': emp['stats']['leave'],
                    'Weekly Off': emp['stats']['weeklyOff'],
                    'Total Days': emp['stats']['total'],
                    'Attendance %': f"{attendance}%"
                })
            
            # Display table
            table_df = pd.DataFrame(table_data)
            st.dataframe(table_df, use_container_width=True, hide_index=True)
            
            # Download Report
            st.subheader("💾 Download Report")
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Summary sheet
                summary_data = {
                    'Metric': ['Total Working Days', 'Days Present', 'Days Absent', 'Leaves Taken', 'Weekly Off Days', 'Overall Attendance %'],
                    'Value': [
                        overall_stats['totalDays'],
                        overall_stats['presentDays'],
                        overall_stats['absentDays'],
                        overall_stats['leaves'],
                        overall_stats['weeklyOff'],
                        (round((overall_stats['presentDays'] / overall_stats['totalDays'] * 100)) if overall_stats['totalDays'] > 0 else 0)
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Employee sheet
                table_df.to_excel(writer, sheet_name='Employee Details', index=False)
            
            output.seek(0)
            
            st.download_button(
                label="📥 Download Excel Report",
                data=output.getvalue(),
                file_name=f"Attendance_Report_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No employee data found")
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.info("Make sure your Excel file has the correct attendance format")

else:
    st.info("👆 Upload an Excel file (.xlsx) to begin")
