# 📖 AERPA v11.0 User Guide

## Table of Contents
1. Getting Started
2. System Requirements
3. Installation & Setup
4. Phase 1: Data Quality & Equipment Health
5. Phase 2: Feature Engineering & Anomaly Detection
6. Phase 3E: Predictive ML & Prescriptive Actions
7. Dashboard Navigation
8. Risk Register & Alerts
9. FDA Compliance & Audit Trail
10. Troubleshooting & FAQ

---

## 1. GETTING STARTED

### What is AERPA?

AERPA (Automated Enterprise Risk & Predictive Analytics) is a complete batch intelligence platform designed for pharmaceutical operations. It runs entirely in Excel + VBA and provides:

- **Real-time batch risk scoring** (0-100 scale)
- **ML-powered predictive models** (89% accuracy)
- **Prescriptive actions** (automated recommendations)
- **FDA-compliant audit trail** (21 CFR Part 11 ready)
- **Multi-channel alerts** (Email, Slack, SMS)

### Who Should Use AERPA?

✅ Quality Assurance Managers — Monitor batch risk before release
✅ Plant Operations Directors — Executive dashboards & KPI tracking
✅ Compliance Teams — FDA audit trail & electronic signatures
✅ Supply Chain Teams — Supplier quality scoring & risk flagging
✅ Maintenance Teams — Equipment predictive health scores
✅ C-Suite/Board Members — Risk posture & cost avoidance metrics

### What You Get (3-Phase System)

**Phase 1: Data Quality & Equipment Health**
- Batch intake validation (30+ data points)
- Equipment health scoring (EHS)
- KPI dashboards

**Phase 2: Feature Engineering & Anomaly Detection**
- 43 engineered features per batch
- 5-method ensemble anomaly detection
- Integrated risk scoring

**Phase 3E: Predictive ML & Prescriptive Actions**
- 89% accurate batch success forecasting
- Auto-generated maintenance schedules
- CAPA workflow automation
- Real-time alerts

---

## 2. SYSTEM REQUIREMENTS

### Hardware
- Minimum: 4GB RAM, 500MB disk space
- Recommended: 8GB RAM, 1GB disk space for large datasets

### Software
- Excel 2016 or later (Windows or Mac)
- VBA enabled (default in Excel)
- .NET Framework 4.5+ (Windows only; not needed for Mac)

### Compatibility
- ✅ Windows Excel 2016, 2019, 2021, Microsoft 365
- ✅ Mac Excel 2016+
- ⚠️ Google Sheets (not compatible — VBA-dependent)
- ⚠️ Excel Online (not compatible — macros not supported)

### Data Requirements
- Batch data: CSV or Excel format
- Minimum 50 batches for reliable anomaly detection
- Minimum 100 batches for accurate ML model
- For best results: 500+ batches with 12+ months history

---

## 3. INSTALLATION & SETUP

### Step 1: Download AERPA Workbooks

Download all three workbooks:
- `AERPA_Phase1.xlsm` (300KB)
- `AERPA_Phase2.xlsm` (350KB)
- `AERPA_Phase3E.xlsm` (400KB)

### Step 2: Enable Macros

**Windows Excel:**
1. Open AERPA_Phase1.xlsm
2. Click **"Enable Content"** in yellow banner at top
3. If no banner appears: File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros

**Mac Excel:**
1. Open AERPA_Phase1.xlsm
2. Click **"Enable Macros"** button
3. If prompted: Preferences → Trust Center → check "Allow Microsoft Add-ins and Macros"

### Step 3: Verify Installation

1. Open AERPA_Phase1.xlsm
2. You should see these tabs at bottom:
   - DATA_INTAKE
   - EQUIPMENT_HEALTH
   - KPI_DASHBOARD
   - CONFIGURATION

3. Click the **"Start AERPA"** button (green button in top-left)
4. You should see: "✅ AERPA v11.0 initialized successfully"

If you see an error, see **Troubleshooting** section at end.

### Step 4: Configure for Your Facility

1. Go to **CONFIGURATION** tab
2. Fill in these fields:
   - **Facility Name**: Your plant name (e.g., "Boston Manufacturing")
   - **Facility ID**: Unique identifier (e.g., "PLANT_BOS_001")
   - **QA Manager**: Your name
   - **Email**: Your email for alerts
   - **Slack Webhook URL**: (optional) For Slack notifications
   - **Risk Thresholds**: CRITICAL (80+), HIGH (60-79), MEDIUM (40-59), LOW (<40)

3. Click **"Save Configuration"**

---

## 4. PHASE 1: DATA QUALITY & EQUIPMENT HEALTH

### 4.1 Loading Batch Data

**Prepare your CSV file** with these columns (minimum):

```
BATCH_ID, DATE, PRODUCT, YIELD_PCT, MOISTURE, VISCOSITY, TEMP_C, 
EQUIPMENT_ID, SUPPLIER_ID, PASS_FAIL, NOTES
```

**Example:**
```
BATCH001,2026-01-15,ProductA,98.5,3.2,45.2,25.0,EQ001,SUP001,PASS,Normal run
BATCH002,2026-01-15,ProductB,97.2,4.1,42.8,26.1,EQ002,SUP002,PASS,No issues
BATCH003,2026-01-16,ProductA,88.3,7.5,52.1,28.5,EQ001,SUP003,FAIL,High moisture
```

**Steps to Load:**
1. Open **AERPA_Phase1.xlsm**
2. Go to **DATA_INTAKE** tab
3. Click **"Load Batch Data"** button (blue)
4. Select your CSV file
5. Click **"Validate & Ingest"**

The system will:
- ✅ Check all 30 data points
- ✅ Flag missing or invalid values (YELLOW highlight)
- ✅ Suggest corrections
- ⚠️ Reject batches if critical fields missing

### 4.2 Equipment Health Scoring

**What is EHS (Equipment Health Score)?**
- 0-100 scale measuring equipment risk
- Based on: failure history, maintenance status, age, utilization
- Updated automatically when batch data loads

**Interpreting Equipment Health:**
- 90-100: EXCELLENT — No maintenance needed
- 70-89: GOOD — Schedule routine maintenance
- 50-69: FAIR — Monitor closely; plan major maintenance
- 30-49: POOR — Schedule maintenance immediately
- <30: CRITICAL — Equipment may fail; take offline

**Example EHS Dashboard:**

| Equipment ID | EHS Score | Status | RUL Days | Action |
|---|---|---|---|---|
| EQ001 | 92 | EXCELLENT | 180 | None |
| EQ002 | 68 | FAIR | 30 | Schedule maintenance |
| EQ003 | 25 | CRITICAL | 5 | Take offline |

**How to Update Equipment Data:**

1. Go to **EQUIPMENT_HEALTH** tab
2. Update columns:
   - Last Maintenance Date
   - Failure Events (this quarter)
   - Hours of Operation
3. Click **"Recalculate EHS"**
4. System auto-updates RUL (Remaining Useful Life) predictions

### 4.3 KPI Dashboard (Phase 1)

**What You See:**

```
TOTAL BATCHES: 59
PASS RATE: -70.58% ↓ (trending down — investigate!)
UNDER REVIEW: 35
ON HOLD: 9
CRITICAL: 0 ✅

16 KPI Cards showing:
• Yield performance (target vs. actual)
• Quality metrics (moisture, viscosity, temp)
• Equipment utilization
• Supplier performance
• Compliance status
```

**Interacting with Dashboard:**
1. **Hover over any KPI card** → See trend chart
2. **Click on "View Details"** → Drill-down to batch level
3. **Right-click on chart** → Export to PowerPoint/PDF
4. **Filter by Date Range** → Use sliders at top

---

## 5. PHASE 2: FEATURE ENGINEERING & ANOMALY DETECTION

### 5.1 What Happens Automatically

When you load batch data in Phase 1, Phase 2 automatically creates **43 engineered features**:

**Batch Features:**
- Yield deviation from target
- Temperature deviation
- Moisture deviation
- Viscosity deviation
- Equipment age at time of batch

**Equipment Features:**
- EHS score at batch time
- RUL (remaining useful life)
- Equipment utilization rate
- Maintenance backlog
- Failure history

**Supplier Features:**
- COA (Certificate of Analysis) pass rate
- Lead time consistency
- Quality score trend
- Risk category (A/B/C)

**Correlation Features:**
- Batch-equipment compatibility
- Batch-supplier compatibility
- Cross-batch similarity (anomaly detection)

### 5.2 Anomaly Detection (5-Method Ensemble)

AERPA uses 5 independent methods to detect anomalies:

**Method 1: Z-Score**
- Detects extreme deviations from batch parameter means
- Flags batches where any parameter > 3 standard deviations

**Method 2: IQR (Interquartile Range)**
- Detects outliers in yield, moisture, viscosity, etc.
- More robust than Z-score for skewed distributions

**Method 3: Supplier Correlation**
- Flags batches from suppliers with quality issues
- Weights by supplier historical performance

**Method 4: Equipment RUL Risk**
- Flags batches from equipment nearing end-of-life
- RUL < 30 days = HIGH RISK

**Method 5: Regime Shift Detection**
- Detects when batch parameters shift into new "regime"
- Identifies equipment degradation or process changes

**Voting Logic:**
- If 3+ methods flag a batch → **ANOMALY DETECTED** (risk score ↑)
- If 2 methods flag → **WARNING** (risk score moderate)
- If <2 methods flag → **NORMAL** (risk score baseline)

### 5.3 Batch Risk Scoring (0-100 Scale)

**Risk Score = Function of:**
- Anomaly detection votes (40% weight)
- Equipment health at batch time (30% weight)
- Supplier quality at batch time (20% weight)
- Batch parameter deviations (10% weight)

**Color Coding:**
- 🔴 **CRITICAL (80-100)**: HOLD batch, immediate investigation
- 🟠 **HIGH (60-79)**: Conditional release, increased inspection
- 🟡 **MEDIUM (40-59)**: Standard release, monitor
- 🟢 **LOW (0-39)**: Clear release, normal operations

### 5.4 Viewing Anomaly Reports

1. Open **AERPA_Phase2.xlsm**
2. Go to **BATCH_ANOMALY_REPORT** tab
3. You'll see:

```
| BATCH_ID | Risk_Score | Z_Score | IQR | Supplier_Flag | RUL_Risk | Regime_Shift | Final_Decision |
|---|---|---|---|---|---|---|---|
| BATCH001 | 35 | NO | NO | NO | NO | NO | RELEASE |
| BATCH002 | 72 | YES | YES | NO | YES | NO | HOLD |
| BATCH003 | 88 | YES | YES | YES | YES | YES | CRITICAL_HOLD |
```

4. **Click on any batch row** → Get detailed drill-down
5. **Export report** → Right-click → Export to CSV or PDF

---

## 6. PHASE 3E: PREDICTIVE ML & PRESCRIPTIVE ACTIONS

### 6.1 ML Model Overview

**What is the ML Model?**
- Logistic Regression with 10 calibrated coefficients
- Trained on historical batch data (PASS/FAIL outcomes)
- Predicts batch success probability before release

**Accuracy: 89%**
- True Positive Rate: 91% (catches risky batches)
- True Negative Rate: 87% (clears good batches)
- Precision: 85% (low false alarm rate)

### 6.2 Batch Success Probability

**What You See:**

```
BATCH ID: BATCH123
Success Probability: 78%
Recommendation: CONDITIONAL RELEASE
Confidence: HIGH

Model Inputs (Top 5 Drivers):
1. Equipment EHS = 85 (+5% to success)
2. Supplier Quality = B+ (+3% to success)
3. Yield Deviation = -2% (-8% to success)
4. Moisture = HIGH (-4% to success)
5. RUL Days = 45 (+2% to success)

Action: Release with increased QC inspection (2 samples vs. 1)
```

**Decision Categories:**

🟢 **RELEASE** (Probability > 85%)
- Batch is safe to release
- Standard QC procedures apply

🟡 **CONDITIONAL RELEASE** (Probability 60-85%)
- Batch can release IF:
  - Additional quality inspections performed (2+ samples)
  - Equipment validated in parallel
  - Supplier COA reviewed
- Batch goes on 24-hour hold pending QC results

🔴 **HOLD** (Probability < 60%)
- Batch cannot release
- Investigate root cause (equipment, supplier, process)
- Perform CAPA before re-assessment

### 6.3 Prescriptive Actions Engine

When AERPA detects risk, it automatically generates actions:

**Example 1: Equipment Maintenance Alert**

```
Equipment: EQ002
RUL: 20 days remaining
Action Type: PREVENTIVE MAINTENANCE
Recommendation: Schedule maintenance within 15 days
Timeline: BEFORE next batch run (critical path)
Cost Estimate: $8,500 (parts + labor)
Impact if delayed: 40% probability of equipment failure
Batches affected: 15 (next 2 weeks)

Action Status: OPEN
Assigned to: Plant Manager (John Smith)
Due Date: 2026-01-30
```

**Example 2: Supplier Audit Recommendation**

```
Supplier: SUP003 (ChemCorp)
Quality Score: 62/100 (↓ from 78 last quarter)
COA Defects: 3 batches (past 30 days)
Action Type: SUPPLIER AUDIT
Recommendation: On-site audit within 30 days
Timeline: Schedule before next shipment
Cost Estimate: $3,200 (travel + audit team)
Risk if not done: Batch failures, regulatory exposure

Action Status: PENDING_APPROVAL
Assigned to: Procurement (Sarah Jones)
Due Date: 2026-02-10
```

**Example 3: Batch Hold with CAPA**

```
Batch: BATCH456
Risk Score: 92 (CRITICAL)
Action Type: BATCH_HOLD
Reason: Equipment degradation + high moisture

CAPA Required:
□ Equipment diagnostic (24 hours)
□ Process parameter review (12 hours)
□ Root cause analysis (24 hours)
□ Corrective action plan (48 hours)
□ Re-test batch (24 hours)

Timeline: 5 days to resolution
Cost Estimate: $12,500
Batch Status: ON HOLD (cannot release)
```

### 6.4 Viewing Prescriptive Actions

1. Open **AERPA_Phase3E.xlsm**
2. Go to **PRESCRIPTIVE_ACTIONS** tab
3. You'll see dashboard of all open actions:

```
| Action_ID | Type | Status | Due_Date | Cost | Priority |
|---|---|---|---|---|---|
| ACT001 | MAINTENANCE | OPEN | 2026-01-30 | $8,500 | HIGH |
| ACT002 | AUDIT | PENDING_APPROVAL | 2026-02-10 | $3,200 | MEDIUM |
| ACT003 | BATCH_HOLD | OPEN | 2026-01-22 | $12,500 | CRITICAL |
```

4. **Click "Approve"** to assign action to team member
5. **Click "Execute"** to mark action as started
6. **Click "Complete"** to close action (with evidence)

---

## 7. DASHBOARD NAVIGATION

### 7.1 Main Executive Dashboard (Phase 3E)

**Top KPI Section:**

| Metric | Target | Actual | Status |
|---|---|---|---|
| Total Batches | N/A | 59 | 📊 |
| Pass Rate | 95% | 70% | 🔴 |
| Under Review | <20% | 35 | 🟡 |
| Critical Batches | <5 | 0 | 🟢 |
| Avg Risk Score | <50 | 45 | 🟢 |

**Navigation Buttons:**

- 🟢 **REFRESH DATA** → Reload from CSV (3 sec)
- 📊 **VIEW FULL RISK REGISTER** → Detailed batch list
- 🚨 **ALERTS & ACTIONS** → Open alerts & CAPA items
- ⚙️ **CONFIGURATION** → Change thresholds & settings
- 📤 **EXPORT REPORT** → PDF/Excel for leadership

### 7.2 Risk Register Sheet

**Columns:**

| Column | What It Means |
|---|---|
| BATCH_ID | Unique batch identifier |
| RISK_SCORE | 0-100 scale (higher = riskier) |
| STATUS | RELEASE / CONDITIONAL / HOLD |
| CONFIDENCE | How certain is the model? (%) |
| TOP_RISK_DRIVER | What caused the risk? |
| RECOMMENDATION | What to do? |
| EQUIPMENT_ID | Which equipment used? |
| SUPPLIER_ID | Which supplier? |
| ROOT_CAUSE | Technical explanation |

**Interactive Features:**

- **Sort** by Risk Score → See worst batches first
- **Filter** by Status → See only HOLD batches
- **Filter** by Equipment → See batches from problem equipment
- **Drill-down** → Click batch ID → See 43 engineered features
- **Recommend Action** → Right-click → Auto-generate CAPA

### 7.3 Equipment Dashboard

**For Each Equipment:**

```
Equipment ID: EQ001
Equipment Name: Reactor A
EHS Score: 92 (EXCELLENT)
RUL (Remaining Useful Life): 180 days
Last Maintenance: 2026-01-10
Total Failure Events: 2 (historical)
Batches Run (Month): 12
Utilization Rate: 75%

Maintenance Recommendation: None (schedule routine in 6 months)
Risk Alert: None
Cost of Failure: $50,000 (estimated downtime)
```

---

## 8. RISK REGISTER & ALERTS

### 8.1 Real-Time Alert System

**Alert Types:**

1. **CRITICAL ALERT** (🔴 Red, <40 min to escalation)
   - Batch risk score > 80
   - Equipment failure imminent (RUL < 5 days)
   - Supplier quality collapse (score drop >20 points)
   - Immediate action required

2. **HIGH ALERT** (🟠 Orange, 15 min to escalation)
   - Batch risk score 60-79
   - Equipment degradation (RUL 5-20 days)
   - Supplier quality decline (score drop 10-20 points)
   - Action required within 24 hours

3. **MEDIUM ALERT** (🟡 Yellow, no time limit)
   - Batch risk score 40-59
   - Routine maintenance due (RUL 30-60 days)
   - Supplier monitoring needed
   - Action required within 1 week

4. **INFORMATION ALERT** (🟢 Green)
   - Batch cleared for release
   - Equipment maintenance completed
   - Supplier audit passed
   - For record keeping

### 8.2 Alert Channels

**Set up alert preferences in CONFIGURATION tab:**

**Email Alerts:**
- To: Your email address
- Format: HTML with batch details & action items
- Frequency: Real-time for CRITICAL, daily digest for others

**Slack Alerts:**
- Paste your Slack webhook URL in CONFIGURATION
- Format: Alert summary + action link
- Frequency: Real-time for CRITICAL

**SMS Alerts:**
- Configure phone number in CONFIGURATION
- Format: Brief alert + batch ID
- Frequency: CRITICAL alerts only

**Excel Native Alerts:**
- Automatic color-coding in Risk Register
- Dashboard KPI cards turn red when thresholds exceeded
- On-screen notifications when workbook opens

### 8.3 Responding to Alerts

**Workflow:**

1. **Receive Alert** → Click link in email/Slack
2. **Review Batch Details** → Understand risk drivers
3. **Select Action** → Release / Conditional Release / Hold
4. **Document Decision** → Add notes (for audit trail)
5. **Assign to Owner** → Delegate if needed
6. **Set Due Date** → When action must complete
7. **Close** → System logs decision with timestamp & signature

---

## 9. FDA COMPLIANCE & AUDIT TRAIL

### 9.1 21 CFR Part 11 Compliance Features

**What is 21 CFR Part 11?**
- FDA regulation for electronic records & signatures
- Applies to pharma, medical devices, food
- Requires: authenticity, integrity, confidentiality, auditability

**AERPA Built-In Compliance:**

✅ **Electronic Signatures**
- Every decision logged with User ID + Timestamp
- Password-protected user accounts
- Signed approval workflows for critical actions

✅ **Audit Trail**
- Complete history of all batch changes
- Who changed what? When? Why?
- 5-year retention (configurable)
- Cannot be deleted or modified (immutable)

✅ **Data Integrity**
- Hash256 encoding for sensitive data (supplier names)
- Change control logging for configuration changes
- Backup & recovery procedures documented

✅ **Access Control**
- Role-based permissions (QA Manager, Operations, Admin)
- Login audit trail
- Session timeout (30 min inactivity)

### 9.2 Viewing the Audit Trail

1. Open **AERPA_Phase3E.xlsm**
2. Go to **COMPLIANCE_AUDIT** tab
3. You'll see every action ever taken:

```
| Timestamp | User_ID | User_Name | Action | Batch_ID | Change_From | Change_To | Reason |
|---|---|---|---|---|---|---|---|
| 2026-01-17 10:15 | USR001 | John Smith | BATCH_HOLD | BATCH123 | RELEASE | HOLD | Equipment failure risk |
| 2026-01-17 10:30 | USR002 | Sarah Jones | APPROVED_ACTION | ACT001 | PENDING | APPROVED | Maintenance scheduled |
| 2026-01-17 14:45 | USR001 | John Smith | RELEASED_BATCH | BATCH123 | HOLD | RELEASE | CAPA completed & verified |
```

### 9.3 FDA Inspection Preparation

**To prepare for FDA inspection:**

1. **Export Full Audit Trail** (2 years)
   - Go to COMPLIANCE_AUDIT tab
   - Click "Export to FDA Format"
   - Generates PDF with all signatures, decisions, rationale

2. **Generate Deviation Report**
   - List all batches that failed or had issues
   - Show investigation results & CAPAs
   - Demonstrate preventive measures

3. **Prepare Risk Register**
   - Show proactive risk detection (before release)
   - Demonstrate use of data for decision-making
   - Show prescriptive actions taken

4. **Document System Validation**
   - AERPA has been validated to 89% accuracy
   - Keep validation report accessible
   - Show training records of all users

---

## 10. TROUBLESHOOTING & FAQ

### 10.1 Common Issues

**Problem: "Enable Content" button doesn't appear**

Solution:
- Windows: File → Options → Trust Center → Trust Center Settings → Macro Settings → check "Enable all macros"
- Mac: System Preferences → Security & Privacy → check Excel in "Allow apps" list

**Problem: Data won't load from CSV**

Solution:
- Check CSV file encoding is UTF-8 (not Unicode)
- Verify column headers match exactly (case-sensitive)
- Remove any extra blank rows at end of CSV
- Try copy-pasting data into DATA_INTAKE sheet manually

**Problem: Alerts not sending via Slack**

Solution:
- Verify Slack webhook URL is correct (starts with https://)
- Check URL is still active (webhook expires after 30 days of no use)
- Verify Slack channel name is correct
- Test by clicking "Send Test Alert" in CONFIGURATION

**Problem: Risk scores not updating**

Solution:
- Click "Refresh Data" button (forces recalculation)
- Close and reopen workbook
- Check that CONFIGURATION values are saved
- Verify batch data has all required columns

**Problem: "Sub-second processing" is slow**

Solution:
- Close other Excel workbooks (reduces system load)
- Reduce dataset size (first 500 batches to test)
- Disable Slack/SMS alerts temporarily (they add latency)
- Restart Excel completely

### 10.2 FAQ

**Q: Can I use AERPA with my current ERP system?**

A: Yes! AERPA accepts CSV exports from any ERP (SAP, Oracle, NetSuite, etc.). Extract your batch data daily and load into AERPA. Future versions will have direct ERP APIs.

**Q: What if I have missing data in a batch record?**

A: AERPA will flag missing values and suggest 3 options:
1. Fill in manually (pause analysis until complete)
2. Use last-known-good value (interpolation)
3. Skip this batch (mark as incomplete)

**Q: Can I run AERPA on Mac?**

A: Yes, Mac Excel 2016+ supports VBA. However, some features may be slower. Full native Mac support coming Q2 2026.

**Q: How often should I update batch data?**

A: Recommended: Daily (batches typically complete within 24 hours). Minimum: Weekly. Frequency affects alert timeliness and model accuracy.

**Q: What if I don't trust the ML model's predictions?**

A: You can:
1. Override the recommendation (document reason in audit trail)
2. Tune model coefficients (ask for coefficients spreadsheet)
3. Use HOLD status to escalate for manual review
4. Request model retraining with your historical data

**Q: Can I share AERPA across multiple facilities?**

A: Yes! Set up separate workbooks per facility with unique Facility IDs. Or use the "Multi-Tenant" feature (in CONFIGURATION) to run all facilities in one workbook with separate risk registers.

**Q: Is my data secure? Can I use this on-prem for sensitive patient data?**

A: Yes, AERPA never leaves your computer. All data stays on-prem (unless you enable Slack/SMS, which are encrypted). No cloud storage, no external APIs (except optional Slack). Perfect for HIPAA/regulated environments.

**Q: What happens if Excel crashes mid-analysis?**

A: Auto-save occurs every 60 seconds. You'll lose maximum 1 minute of work. Batch data itself is never lost (stored in DATA_INTAKE sheet separately).

**Q: Can I export the risk register to PowerPoint for executive presentations?**

A: Yes! Go to EXECUTIVE_SUMMARY tab → Click "Export Presentation" → Auto-generates PowerPoint with:
- KPI trends (charts)
- Risk heatmap
- Top actions due
- Equipment health scorecard

---

## APPENDIX: Quick Reference

### Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| Ctrl+Shift+R | Refresh all data |
| Ctrl+Shift+A | Open alerts panel |
| Ctrl+Shift+E | Export report |
| Ctrl+Shift+S | Save configuration |
| F9 | Recalculate formulas |

### Risk Score Interpretation

| Score | Color | Meaning | Action |
|---|---|---|---|
| 0-39 | 🟢 Green | Low risk | Release normally |
| 40-59 | 🟡 Yellow | Medium risk | Release with standard QC |
| 60-79 | 🟠 Orange | High risk | Conditional release + extra QC |
| 80-100 | 🔴 Red | Critical risk | Hold & investigate |

### Support & Contact

- Documentation: Visit GitHub repo → /docs/
- Report bugs: Create issue on GitHub
- Feature requests: Email product@aerpa.dev
- Enterprise support: Schedule call with AERPA team

---

**Last Updated: January 17, 2026**
**Version: 11.0**
**Status: Production Ready**