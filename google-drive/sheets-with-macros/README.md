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

## Templates (17 Total)

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

#### 💵 Budget Planning with Scenario Modeling ⭐ NEW
**Files:** `budget-planning.csv` + `budget-planning.gs`
- Multiple budget scenarios (Best/Base/Worst case)
- Revenue forecasting with growth models
- Cash flow projections (12-month)
- Break-even analysis
- Variance analysis (Actual vs Budget)
- Department budgets, quarterly rollups
- Startup runway calculator

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

#### 📅 Meeting Scheduler with Calendar ⭐ NEW
**Files:** `meeting-scheduler.csv` + `meeting-scheduler.gs`
- Create calendar events directly from sheet
- Recurring meeting templates
- Attendee management, availability checking
- Meeting templates (1:1, Standup, Sprint, Board)
- Meeting notes and action items
- Meeting cost calculator
- Analytics and reporting

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

#### 🏢 Vendor Scoring & Management ⭐ NEW
**Files:** `vendor-scoring.csv` + `vendor-scoring.gs`
- Vendor evaluation scorecards
- Weighted criteria scoring (7 criteria)
- RFP/RFI generation
- Performance monitoring, SLA tracking
- Risk assessment, compliance verification
- Vendor comparison reports
- Renewal alerts

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

### Productivity & Organization

#### 📁 Google Drive Organizer ⭐ NEW
**Files:** `drive-organizer.csv` + `drive-organizer.gs`
- Scan entire Drive for all files
- Auto-categorize by file type and keywords
- Duplicate file detection
- Create BlackRoad folder structure (29 folders)
- Batch move files to organized folders
- Archive old files (1+ year)
- Storage analytics and reporting

---

## Automation Triggers

All templates support automatic scheduling:

1. Go to **Extensions > Apps Script**
2. Click the **clock icon** (Triggers)
3. Click **+ Add Trigger**
4. Select function (e.g., `refreshAllData`, `dailyLowStockCheck`, `checkComplianceAlerts`)
5. Choose time-based trigger
6. Set frequency (daily/weekly)

## BlackRoad Folder Structure

The Drive Organizer creates this structure:

```
BlackRoad OS/
├── Corporate/
│   ├── Formation
│   ├── Legal
│   ├── Tax
│   └── Compliance
├── Finance/
│   ├── Invoices
│   ├── Expenses
│   └── Reports
├── HR/
│   ├── Recruiting
│   ├── Onboarding
│   └── Policies
├── Engineering/
│   ├── Architecture
│   ├── Documentation
│   └── Specs
├── Marketing/
│   ├── Pitch Decks
│   ├── Whitepapers
│   └── Brand
├── Sales/
│   ├── Proposals
│   ├── Contracts
│   └── Pipeline
├── Products/
│   ├── Prism Console
│   ├── Agent Swarm
│   └── Documentation
├── Templates/
│   ├── Sheets
│   ├── Docs
│   └── Slides
├── Archive/
│   ├── 2024
│   └── 2023
└── Personal/
    ├── Resumes
    └── Notes
```

## Security Notes

- Scripts run with **your permissions**
- Email sending uses **your Gmail**
- Grant permissions when prompted
- Review code before running

## Customization

Edit CONFIG sections in each script:

```javascript
const CONFIG = {
  COMPANY_NAME: 'BlackRoad OS, Inc.',
  SENDER_NAME: 'Your Name',
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

**Calendar events not syncing?**
- Ensure calendar permissions granted
- Check timezone settings in CONFIG

---

## Template Summary

| # | Template | Menu | Key Features |
|---|----------|------|--------------|
| 1 | Invoice Generator | 📄 Invoice | Auto-numbering, PDF email |
| 2 | Expense Tracker | 💰 Expenses | Approval workflow, mileage |
| 3 | Financial Dashboard | 📊 Finance | KPIs, bank import |
| 4 | Sales Pipeline | 💼 Sales | Forecasting, velocity |
| 5 | Budget Planning | 💵 Budget | Scenarios, runway calc |
| 6 | Time Tracking | ⏰ Time | Clock in/out, overtime |
| 7 | HR Onboarding | 👥 HR | 17-task checklist |
| 8 | CRM Automation | 🎯 CRM | Lead scoring, sequences |
| 9 | Meeting Scheduler | 📅 Meetings | Calendar sync, templates |
| 10 | Project Management | 📈 Projects | Gantt, dependencies |
| 11 | Inventory Management | 📦 Inventory | SKU lookup, PO generation |
| 12 | Contract Management | 📝 Contracts | Lifecycle, renewals |
| 13 | Vendor Scoring | 🏢 Vendors | Scorecards, RFP |
| 14 | HIPAA Compliance | 🏥 HIPAA | PHI logging, BAAs |
| 15 | SOX Compliance | 📈 SOX | Control testing |
| 16 | GDPR Compliance | 🇪🇺 GDPR | DSR tracking |
| 17 | Drive Organizer | 📁 Drive | File organization |

---

*Generated by BlackRoad OS, Inc.*
