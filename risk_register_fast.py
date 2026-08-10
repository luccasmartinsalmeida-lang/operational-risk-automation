import os
import sys
import pandas as pd

df = pd.read_excel('risk_register_practice_data.xlsx')

# Normalizes header case so 'Probability', 'PROBABILITY' or 'probability' don't crash the script
df.columns = df.columns.str.strip().str.lower()

# Defensive check for required calculation columns
if not {'probability', 'impact'}.issubset(df.columns):
    raise KeyError("Missing required 'probability' or 'impact' columns in input file")

# Clean text whitespace
str_cols = df.select_dtypes(include='object').columns
df[str_cols] = df[str_cols].apply(lambda x: x.str.strip())

# Fix date format
if 'date_identified' in df.columns:
    df['date_identified'] = pd.to_datetime(df['date_identified'], errors='coerce').dt.strftime('%Y-%m-%d')

# Core calculations
df['risk_score'] = df['probability'] * df['impact']
df['severity'] = pd.cut(
    df['risk_score'],
    bins=[-1, 4, 9, 14, 100],
    labels=['Low', 'Medium', 'High', 'Critical']
)

df = df.sort_values(by=['risk_score', 'probability'], ascending=False).reset_index(drop=True)

cols = {
    'id': 'Risk ID', 'risk_name': 'Risk Name', 'category': 'Category',
    'department': 'Department', 'owner': 'Owner', 'probability': 'Probability',
    'impact': 'Impact', 'risk_score': 'Risk Score', 'severity': 'Severity',
    'status': 'Status', 'date_identified': 'Date Identified', 'description': 'Description'
}
df = df.rename(columns=cols)

# Build summary with forced ordering
summary = df['Severity'].value_counts().reindex(['Critical', 'High', 'Medium', 'Low']).fillna(0).astype(int).to_frame('Count')
top_risks = df[(df['Severity'].isin(['Critical', 'High'])) & (df['Status'] == 'Open')]

out = 'Operational_Risk_Register.xlsx'
with pd.ExcelWriter(out) as writer:
    summary.to_excel(writer, sheet_name='Summary')
    df.to_excel(writer, sheet_name='Risk_Register', index=False)
    top_risks.to_excel(writer, sheet_name='Top_Priority_Gaps', index=False)

if sys.platform == 'win32':
    os.startfile(out)
elif sys.platform == 'darwin':
    os.system(f'open "{out}"')