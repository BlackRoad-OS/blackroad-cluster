# BlackRoad OS - Google Sheets with Macros

Enterprise-grade spreadsheet templates with Google Apps Script automation.

## Setup Instructions

1. **Import the CSV** to Google Sheets
2. Go to **Extensions > Apps Script**
3. Delete any existing code
4. **Paste the corresponding .gs file** contents
5. Click **Save** (Ctrl+S)
6. **Refresh** your Google Sheet
7. Look for the new **custom menu** in the menu bar

## Templates (13 Total)

### Business Operations

#### 📄 Invoice Generator
**Files:** `invoice-generator.csv` + `invoice-generator.gs`
- Auto-increment invoice numbers
- Calculate due dates from payment terms
- Send invoices via Gmail as PDF
- Track invoice status (Draft/Sent/Paid/Overdue)
- Generate monthly reports, overdue alerts

#### 💰 Expense Tracker
**Files:** `expense-tracker.csv` + `expense-tracker.gs`
- Quick add expense dialog
- Attach receipts from Google Drive
- Approval workflow (Approve/Reject)
- Mileage calculator (IRS rate $0.67/mi)
- Per diem calculator (GSA rates)
- Export for QuickBooks/Xero

#### 📊 Financial Dashboard
**Files:** `financial-dashboard.csv` + `financial-dashboard.gs`
- KPI cards with trend analysis
- Import bank CSV statements
- Cash flow forecasting, AR aging
- Budget vs actual tracking
- Auto-refresh triggers (daily/weekly)

#### 💼 Sales Pipeline
**Files:** `sales-pipeline.csv` + `sales-pipeline.gs`
- Visual pipeline stages with probability weighting
- Revenue forecasting (weighted/unweighted)
- Sales velocity metrics
- Rep performance dashboards
- Win/loss analysis, stalled deal alerts

---

### HR & People

#### ⏰ Time Tracking with Payroll
**Files:** `time-tracking.csv` + `time-tracking.gs`
- Clock in/out with timestamps
- Break time tracking
- Overtime calculations (40hr weekly, 8hr daily)
- Double-time support (12+ hrs/day)
- PTO/sick time requests
- Payroll export

#### 👥 HR Onboarding Workflow
**Files:** `hr-onboarding.csv` + `hr-onboarding.gs`
- 17-task checklist automation
- Individual checklist sheets per employee
- Welcome email sequences
- 30/60/90 day review reminders
- Manager notifications

#### 🎯 CRM with Email Automation
**Files:** `crm-automation.csv` + `crm-automation.gs`
- Contact management with lead scoring
- Email templates with merge fields
- Automated follow-up sequences
- Pipeline reporting, activity logging

---

### Project & Inventory

#### 📈 Project Management with Gantt
**Files:** `project-management.csv` + `project-management.gs`
- Visual Gantt chart auto-generation
- Task dependency tracking
- Resource allocation, milestone alerts
- Progress tracking, status emails
- PDF export

#### 📦 Inventory Management
**Files:** `inventory-management.csv` + `inventory-management.gs`
- SKU/Barcode lookup
- Stock in/out with history
- Low stock alerts, reorder points
- Purchase order generation
- ABC analysis, inventory valuation

#### 📝 Contract Management
**Files:** `contract-management.csv` + `contract-management.gs`
- Contract lifecycle tracking
- Renewal/expiration alerts (60-day notice)
- E-signature status monitoring
- Amendment management
- Approval workflow, value tracking

---

### Compliance

#### 🏥 HIPAA Compliance
**Files:** `hipaa-compliance.csv` + `hipaa-compliance.gs`
- PHI access logging (Article 15)
- Business Associate Agreement tracking
- Security incident management
- Breach notification workflow (72-hour)
- Training compliance monitoring
- Annual audit checklists

#### 📈 SOX Compliance
**Files:** `sox-compliance.csv` + `sox-compliance.gs`
- Control testing automation
- Deficiency management (SD/MW tracking)
- Evidence collection
- Segregation of duties matrix
- Management certification workflow
- Quarter/year-end close checklists

#### 🇪🇺 GDPR Compliance
**Files:** `gdpr-compliance.csv` + `gdpr-compliance.gs`
- Data Subject Request (DSR) tracking
- Processing activities register (Article 30)
- Data breach notification (72-hour DPA, individual)
- Consent management
- DPIA templates
- Cross-border transfer tracking

---

## Automation Triggers

All templates support automatic scheduling:

1. Go to **Extensions > Apps Script**
2. Click the **clock icon** (Triggers)
3. Click **+ Add Trigger**
4. Select function (e.g., `refreshAllData`, `dailyLowStockCheck`, `checkComplianceAlerts`)
5. Choose time-based trigger
6. Set frequency (daily/weekly)

## Security Notes

- Scripts run with **your permissions**
- Email sending uses **your Gmail**
- Grant permissions when prompted
- Review code before running

## Customization

Edit CONFIG sections in each script:

```javascript
const CONFIG = {
  SENDER_NAME: 'Your Name',
  COMPANY_NAME: 'Your Company',
  // ... other settings
};
```

## Troubleshooting

**Menu not appearing?**
- Refresh the page
- Check Extensions > Apps Script for errors

**Permissions error?**
- Click through authorization prompts
- Check your Google account permissions

**Email not sending?**
- Check daily Gmail sending limits (500/day)
- Verify recipient email addresses

---

*Generated by BlackRoad OS, Inc.*
