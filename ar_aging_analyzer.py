import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns

# ===================== CONFIG =====================
DATA_FILE = "ar_data.csv"
OUTPUT_EXCEL = "AR_Aging_Report.xlsx"
CURRENT_DATE = pd.Timestamp.today()   # You can hardcode: pd.Timestamp("2026-04-21")

# ===================== LOAD & PREPARE DATA =====================
print("🔄 Loading AR data...")
df = pd.read_csv(DATA_FILE)

# Convert dates
df['Invoice_Date'] = pd.to_datetime(df['Invoice_Date'])
df['Due_Date'] = pd.to_datetime(df['Due_Date'])

# Calculate days overdue
df['Days_Overdue'] = (CURRENT_DATE - df['Due_Date']).dt.days

# ===================== AGING BUCKETS =====================
def aging_bucket(days):
    if days <= 0:
        return "Current"
    elif days <= 30:
        return "1-30 Days"
    elif days <= 60:
        return "31-60 Days"
    elif days <= 90:
        return "61-90 Days"
    else:
        return "90+ Days"

df['Aging_Bucket'] = df['Days_Overdue'].apply(aging_bucket)

# ===================== SUMMARY REPORT =====================
print("\n" + "="*60)
print("📊 ACCOUNTS RECEIVABLE AGING REPORT")
print("="*60)
print(f"Report Date : {CURRENT_DATE.strftime('%Y-%m-%d')}")
print(f"Total Invoices : {len(df):,}")
print(f"Total Outstanding : ${df['Amount'].sum():,.2f}\n")

summary = df.groupby('Aging_Bucket')['Amount'].agg(['sum', 'count']).reset_index()
summary.columns = ['Aging Bucket', 'Amount', 'Invoice Count']
summary['% of Total'] = (summary['Amount'] / summary['Amount'].sum() * 100).round(1)
summary = summary[['Aging Bucket', 'Invoice Count', 'Amount', '% of Total']]

print(summary.to_string(index=False))

# Save detailed + summary to Excel
with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Detailed_AR', index=False)
    summary.to_excel(writer, sheet_name='Summary', index=False)
print(f"\n💾 Report saved as {OUTPUT_EXCEL}")

# ===================== VISUALIZATIONS =====================
sns.set_style("whitegrid")
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# Bar chart
sns.barplot(data=summary, x='Aging Bucket', y='Amount', ax=axes[0], palette="Reds_r")
axes[0].set_title('AR Aging by Bucket (Amount)')
axes[0].set_ylabel('Amount ($)')
for container in axes[0].containers:
    axes[0].bar_label(container, fmt='$%.0f', label_type='edge')

# Pie chart
axes[1].pie(summary['Amount'], labels=summary['Aging Bucket'], autopct='%1.1f%%',
            colors=sns.color_palette("Reds_r"), startangle=90)
axes[1].set_title('AR Aging Distribution')

plt.tight_layout()
plt.show()

print("\n✅ Analysis complete! Open the Excel file and charts.")
