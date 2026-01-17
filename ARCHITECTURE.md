# AERPA Technical Architecture

## System Overview
AERPA is a 3-phase intelligence platform running 100% on-premises in Excel + VBA.

### Phase 1: Data Quality & Equipment Health
- Input: Raw batch CSV data
- Processing: 30-point validation, equipment health scoring
- Output: Validated data + KPI dashboard

### Phase 2: Feature Engineering & Anomaly Detection
- 43 engineered features (batch parameters, equipment correlation, supplier risk)
- 5-method ensemble anomaly detection
- Risk scoring (0-100 scale)

### Phase 3: ML, Prescriptive Actions & Compliance
- Logistic Regression model (10 coefficients)
- Batch success probability forecasting
- Auto-generated prescriptive actions
- FDA audit trail with electronic signatures

## Dependencies
- Excel 2016+ with VBA enabled
- 0 external libraries (pure VBA)
- Compatible with: Windows Excel, Mac Excel (with limitations)

## Performance
- Sub-second processing for 100+ batches
- In-memory execution (no database needed)
