import pandas as pd
import xlsxwriter

# 1. Load source dataset
filepath = 'risk_register_practice_data.xlsx'
df = pd.read_excel(filepath)

# Clean whitespace from text columns
for col in df.select_dtypes(include=['object']).columns:
    df[col] = df[col].astype(str).str.strip()

# 2. Automate Risk Calculations
df['risk_score'] = df['probability'] * df['impact']

def classify_severity(score):
    if score >= 15:
        return 'Critical'
    elif score >= 10:
        return 'High'
    elif score >= 5:
        return 'Medium'
    else:
        return 'Low'

df['severity'] = df['risk_score'].apply(classify_severity)

# Sort by Risk Score and Probability descending
df = df.sort_values(by=['risk_score', 'probability'], ascending=[False, False]).reset_index(drop=True)

# Rename columns for executive presentation
column_mapping = {
    'id': 'Risk ID',
    'risk_name': 'Risk Name',
    'category': 'Category',
    'department': 'Department',
    'owner': 'Owner',
    'probability': 'Probability',
    'impact': 'Impact',
    'risk_score': 'Risk Score',
    'severity': 'Severity',
    'status': 'Status',
    'date_identified': 'Date Identified',
    'description': 'Description'
}
df = df.rename(columns=column_mapping)

# Reorder columns logically
cols_order = [
    'Risk ID', 'Risk Name', 'Category', 'Department', 'Owner', 
    'Probability', 'Impact', 'Risk Score', 'Severity', 'Status', 
    'Date Identified', 'Description'
]
df = df[cols_order]

# 3. Export Formatted Workbook via XlsxWriter
output_filename = 'Automated_Operational_Risk_Register.xlsx'

with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
    workbook = writer.book

    # Formatting Styles
    header_fmt = workbook.add_format({
        'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    kpi_title_fmt = workbook.add_format({
        'bold': True, 'font_color': '#1F4E78', 'font_size': 11, 'align': 'center'
    })
    kpi_num_fmt = workbook.add_format({
        'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'font_color': '#1F4E78'
    })
    
    # Severity Color Matrix Formats
    critical_fmt = workbook.add_format({'bg_color': '#C00000', 'font_color': 'white', 'bold': True, 'align': 'center'})
    high_fmt = workbook.add_format({'bg_color': '#FFC000', 'font_color': 'black', 'bold': True, 'align': 'center'})
    medium_fmt = workbook.add_format({'bg_color': '#FFF2CC', 'font_color': 'black', 'align': 'center'})
    low_fmt = workbook.add_format({'bg_color': '#E2EFDA', 'font_color': '#375623', 'align': 'center'})

    # --- TAB 1: EXECUTIVE DASHBOARD ---
    dashboard_ws = workbook.add_worksheet('Risk Dashboard')
    dashboard_ws.hide_gridlines(2)

    dashboard_ws.merge_range('A1:F1', 'Operational & IT Risk Management Dashboard', workbook.add_format({
        'bold': True, 'font_size': 16, 'font_color': '#1F4E78', 'valign': 'vcenter'
    }))

    # Compute Dashboard Indicators
    total_risks = len(df)
    critical_count = len(df[df['Severity'] == 'Critical'])
    open_count = len(df[df['Status'] == 'Open'])
    avg_score = df['Risk Score'].mean()

    # Lay out KPI Cards
    dashboard_ws.merge_range('A3:B3', 'Total Identified Risks', kpi_title_fmt)
    dashboard_ws.merge_range('A4:B5', total_risks, kpi_num_fmt)

    dashboard_ws.merge_range('C3:D3', 'Critical Severity Risks', kpi_title_fmt)
    dashboard_ws.merge_range('C4:D5', critical_count, kpi_num_fmt)

    dashboard_ws.merge_range('E3:F3', 'Open Status Risks', kpi_title_fmt)
    dashboard_ws.merge_range('E4:F5', open_count, kpi_num_fmt)

    dashboard_ws.merge_range('G3:H3', 'Average Risk Score', kpi_title_fmt)
    dashboard_ws.merge_range('G4:H5', f"{avg_score:.1f}", kpi_num_fmt)

    # Severity Summary Source Table
    dashboard_ws.write('A8', 'Severity Level', header_fmt)
    dashboard_ws.write('B8', 'Count', header_fmt)
    
    severity_order = ['Critical', 'High', 'Medium', 'Low']
    for idx, lev in enumerate(severity_order):
        count = len(df[df['Severity'] == lev])
        dashboard_ws.write(f'A{9+idx}', lev)
        dashboard_ws.write(f'B{9+idx}', count)

    # Embed Doughnut Chart
    donut_chart = workbook.add_chart({'type': 'doughnut'})
    donut_chart.add_series({
        'name': 'Risk Severity Distribution',
        'categories': "='Risk Dashboard'!$A$9:$A$12",
        'values': "='Risk Dashboard'!$B$9:$B$12",
        'points': [
            {'fill': {'color': '#C00000'}},  # Critical
            {'fill': {'color': '#FFC000'}},  # High
            {'fill': {'color': '#FFF2CC'}},  # Medium
            {'fill': {'color': '#E2EFDA'}},  # Low
        ]
    })
    donut_chart.set_title({'name': 'Risks by Severity Level'})
    dashboard_ws.insert_chart('A14', donut_chart, {'x_scale': 1.0, 'y_scale': 0.95})

    # --- TAB 2: FULL RISK REGISTER ---
    df.to_excel(writer, sheet_name='Risk Register', index=False)
    register_ws = writer.sheets['Risk Register']

    for col_num, val in enumerate(df.columns.values):
        register_ws.write(0, col_num, val, header_fmt)

    # Apply conditional colors to Severity column
    sev_col_idx = df.columns.get_loc('Severity')
    for r in range(len(df)):
        val = df.iloc[r, sev_col_idx]
        cell_ref = xlsxwriter.utility.xl_rowcol_to_cell(r + 1, sev_col_idx)
        if val == 'Critical':
            register_ws.write(cell_ref, val, critical_fmt)
        elif val == 'High':
            register_ws.write(cell_ref, val, high_fmt)
        elif val == 'Medium':
            register_ws.write(cell_ref, val, medium_fmt)
        elif val == 'Low':
            register_ws.write(cell_ref, val, low_fmt)

    # --- TAB 3: TOP PRIORITY OPEN RISKS ---
    top_df = df[(df['Severity'].isin(['Critical', 'High'])) & (df['Status'] == 'Open')]
    top_df.to_excel(writer, sheet_name='Top Priority Open Risks', index=False)
    top_ws = writer.sheets['Top Priority Open Risks']

    for col_num, val in enumerate(top_df.columns.values):
        top_ws.write(0, col_num, val, header_fmt)

    print(f"Successfully processed {total_risks} risks and generated '{output_filename}'.")