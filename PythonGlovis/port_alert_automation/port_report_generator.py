import pandas as pd 
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from dotenv import load_dotenv
load_dotenv()

# ---------- 0. Setup ----------
date_str = datetime.today().strftime('%m%d%y')
excel_filename = f"weekly_alert_report_{date_str}.xlsx"
csv_filename = f"alert_summary_{date_str}.csv"

# ---------- 1. Load Data from Excel ----------
df = pd.read_excel("port_alert_automation/July-Dummy-Data01.xlsx")
df.columns = df.columns.str.strip()

df['Tender Date'] = pd.to_datetime(df['Tender Date'], errors='coerce')
df['Shipment Date'] = pd.to_datetime(df['Shipment Date'], errors='coerce')

df.dropna(subset=['Dept. Port', '1st Leg Carrier'], inplace=True)
df.rename(columns={'Dept. Port': 'PORT', '1st Leg Carrier': 'Carrier'}, inplace=True)

# ---------- 2. Define Week Dates (hardcoded to 7/14–7/20) ----------
this_monday = datetime(2025, 7, 14)
last_monday = this_monday - pd.Timedelta(weeks=1)

df['WeekStart_Tender'] = df['Tender Date'].dt.to_period('W').dt.start_time
df['WeekStart_Shipment'] = df['Shipment Date'].dt.to_period('W').dt.start_time

# This week
this_week = df[(df['WeekStart_Tender'] == this_monday) | (df['WeekStart_Shipment'] == this_monday)].copy()
this_week['MovementType'] = None
this_week.loc[this_week['WeekStart_Tender'] == this_monday, 'MovementType'] = 'IN'
this_week.loc[this_week['WeekStart_Shipment'] == this_monday, 'MovementType'] = this_week['MovementType'].fillna('OUT')

# Last week
last_week = df[(df['WeekStart_Tender'] == last_monday) | (df['WeekStart_Shipment'] == last_monday)].copy()
last_week['MovementType'] = None
last_week.loc[last_week['WeekStart_Tender'] == last_monday, 'MovementType'] = 'IN'
last_week.loc[last_week['WeekStart_Shipment'] == last_monday, 'MovementType'] = last_week['MovementType'].fillna('OUT')

# ---------- 3. Summarize ----------
def summarize(df):
    df_in = df[df['MovementType'] == 'IN'].groupby(['PORT', 'Carrier']).size().rename('Units In')
    df_out = df[df['MovementType'] == 'OUT'].groupby(['PORT', 'Carrier']).size().rename('Units Out')
    summary = pd.concat([df_in, df_out], axis=1).fillna(0).reset_index()
    summary['Ratio'] = summary['Units Out'] / summary['Units In'].replace(0, pd.NA)
    return summary

summary_this = summarize(this_week)
summary_last = summarize(last_week)

# ---------- 4. Merge & Alert Logic ----------
merged = pd.merge(
    summary_this.groupby('PORT').agg({'Units In':'sum','Units Out':'sum'}).reset_index(),
    summary_last.groupby('PORT').agg({'Units In':'sum','Units Out':'sum'}).reset_index(),
    on='PORT',
    suffixes=('_this', '_last')
)
merged['Ratio_this'] = merged['Units Out_this'] / merged['Units In_this'].replace(0, pd.NA)
merged['Ratio_last'] = merged['Units Out_last'] / merged['Units In_last'].replace(0, pd.NA)
merged['Delta'] = merged['Ratio_this'] - merged['Ratio_last']
alerted = merged[merged['Ratio_this'] < 0.8].copy()

# ---------- 5. Export Excel ----------
if not alerted.empty:
    with pd.ExcelWriter(excel_filename) as writer:
        for port in alerted['PORT']:
            this_week[this_week['PORT'] == port].to_excel(writer, sheet_name=port[:31], index=False)
else:
    print("⚠️ No alerts to export to Excel.")

# ---------- 6. Build Email HTML ----------
html = f"""
<h2>🎉 Weekly Port Alert Report – {datetime.today().strftime('%m/%d/%Y')}</h2>
<p>To reduce manual work and identify inefficiencies early, I built a Python-based alerting tool. 
It automatically calculates Outflow:Inflow ratios, flags low performers, and sends a styled email report every week.</p>
<p><em>This dummy email is auto-generated weekly in HTML + Excel format. It flags operational issues by port and carrier using Python + pandas logic.</em></p>
"""

if alerted.empty:
    html += "<p><i>No ports triggered alerts this week. No inefficiencies detected 🎉</i></p>"
else:
    for port in alerted['PORT']:
        row = merged[merged['PORT'] == port].iloc[0]
        arrow = "🟢⬆️" if row['Delta'] >= 0 else "<span style='color:red;'>🔻</span>"
        html += f"""
        <h3>Dept. Port: <span style='background-color:#DCE6F1; padding:2px 6px;'>{port}</span></h3>
        <p><b>This Week Ratio:</b> {arrow} {row['Ratio_this']:.2f} |
           <b>Last Week Ratio:</b> {row['Ratio_last']:.2f} |
           <b>Delta:</b> {row['Delta']:.2f}</p>
        <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;">
            <tr style="background-color: #f2f2f2;">
                <th>Carrier</th>
                <th>Units In (This)</th>
                <th>Units Out (This)</th>
                <th>Ratio (This)</th>
                <th>Units In (Last)</th>
                <th>Units Out (Last)</th>
                <th>Ratio (Last)</th>
                <th>Delta</th>
            </tr>
        """
        for carrier in summary_this[summary_this['PORT'] == port]['Carrier'].unique():
            this_row = summary_this[(summary_this['PORT'] == port) & (summary_this['Carrier'] == carrier)].iloc[0]
            last_row = summary_last[(summary_last['PORT'] == port) & (summary_last['Carrier'] == carrier)]
            last_row = last_row.iloc[0] if not last_row.empty else {'Units In': 0, 'Units Out': 0}

            in_this, out_this = this_row['Units In'], this_row['Units Out']
            in_last, out_last = last_row['Units In'], last_row['Units Out']

            ratio_this = out_this / in_this if in_this else 0
            ratio_last = out_last / in_last if in_last else 0
            delta = ratio_this - ratio_last
            icon = "🔻" if delta < 0 else "🟢"

            html += f"""
            <tr>
                <td>{carrier}</td>
                <td>{in_this}</td>
                <td>{out_this}</td>
                <td>{ratio_this:.2f}</td>
                <td>{in_last}</td>
                <td>{out_last}</td>
                <td>{ratio_last:.2f}</td>
                <td>{icon} {delta:.2f}</td>
            </tr>
            """
        html += "</table><br><br>"

# ---------- 6b. Signature ----------
html += """
<p>Best Regards,<br>
<strong>Your Name</strong><br>
Data Analyst Intern<br>
Your Company<br><br>
E-mail: <a href="mailto:example@example.com">example@example.com</a><br>
Web: <a href="#">companywebsite.com</a><br>
Company Address
</p>

<p style="font-size:12px; color:gray;">Generated in &lt;2s, delivered to manager inbox before 9am every Monday.</p>
"""

# ---------- 7. Export CSV ----------
alerted.to_csv(csv_filename, index=False)

# ---------- 8. Send Email ----------
SEND_EMAIL = True

if SEND_EMAIL:
    sender_email = os.getenv("SENDER_EMAIL")
    receiver_email = os.getenv("RECEIVER_EMAIL")
    password = os.getenv("EMAIL_PASSWORD")


    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f"⚠️ Weekly Port Alert Report – {datetime.today().strftime('%m/%d/%Y')}"
    msg.attach(MIMEText(html, 'html'))

    if not alerted.empty:
        with open(excel_filename, 'rb') as f:
            part = MIMEApplication(f.read(), _subtype='xlsx')
            part.add_header('Content-Disposition', 'attachment', filename=excel_filename)
            msg.attach(part)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, password)
        smtp.send_message(msg)

    print("✅ Email sent.")
else:
    print("📋 Email preview:")
    print(html)
