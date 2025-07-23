✍️ Author
Jacob Lee – Data Analyst Intern @ Hyundai Glovis

Note: This is a dummy project for demo purposes using anonymized test data.

🚢 Port Alert Automation – Weekly Shipment Report
This Python tool analyzes weekly shipment data to detect inefficiencies by port and carrier. It automatically:

Calculates Outflow:Inflow ratios
Flags ports with low performance (<0.80 ratio)
Exports a formatted Excel report (1 sheet per flagged port)
Generates an HTML email with delta visuals (🔻/⬆️)
Sends the report to a specified email address
📁 Input
July-Dummy-Data.xlsx: Test data simulating weekly shipment counts by port and carrier
📤 Output
weekly_alert_report_<mmddyy>.xlsx: Excel report with VIN-level breakdown
alert_summary_<mmddyy>.csv: Port-level summary with flagged ratios
HTML email preview or actual email sent via SMTP
📦 Dependencies
See requirements.txt

🛠️ Usage
Update port_report_generator.py if using real data
Run the script:
python port_report_generator.py
Email will be sent if SEND_EMAIL = True
