# ETL Automation for Bulk Shipping Document Generation

## Overview
Automated ETL pipeline for bulk generation of shipping documents (DDT)
in a regulated operational environment. Designed to replace a 
critical manual process prone to transcription errors on fiscal codes 
and delivery addresses.

## Business Problem
Manual generation of 58 DDTs per batch required ~45 minutes with a 
5% error rate on critical fields (VAT numbers, delivery addresses), 
impacting daily shipping schedules and fiscal compliance.

## Methodology
Six Sigma DMAIC — Yellow Belt  
Process improved from **3.1 Sigma** (manual) to **6 Sigma** (automated).

## Results
| Metric | Before | After |
|--------|--------|-------|
| Cycle time per batch | ~45 min | ~17 sec |
| Error rate | 5% | 0% |
| Manual intervention | High | Minimal (supervision only) |
| Sigma level | 3.1σ | 6σ |
| Est. annual productivity gain | — | €1,200 |

## Key Insight — Analyze Phase
The Power Query script is nearly instantaneous (0.91s).  
The real bottleneck was **ERP validation and printing (16.07s)**.  
This finding redirected optimization toward system architecture 
rather than input speed — a result only visible through 
precise per-segment measurement.

## Architecture
```
Excel Report (daily)
      ↓
Power Query ETL
  → Data cleaning & normalization
  → VAT/fiscal code Dual-Matching
  → Left Outer Join with master registry
      ↓
VBA Middleware
  → Hierarchical XML generation
  → Dynamic aggregation by client
      ↓
ERP Import (XML)
  → Bulk DDT validation
  → Batch printing
```

## Technical Components
- **Power Query** — extraction, cleaning, normalization  
  (Trim/Clean functions, string standardization, length validation)
- **VBA** — XML generation and ERP integration middleware
- **Dual-Matching logic** — simultaneous VAT + fiscal code lookup
- **Bridge System** — 3-tier architecture for zero record loss  
  (Transactional → Mapping Table → Master)

## Data Quality Controls
- VAT number format validation (11 chars)
- Fiscal code format validation (16 chars)  
- Removal of special characters and invisible spaces
- Text format enforcement to preserve leading zeros
- 100% automated matching for mapped records

## Tools
Power Query · VBA · XML · Excel · ERP Integration

## Documentation
Full Six Sigma DMAIC case study available in the  
[Notion Portfolio](https://www.notion.so/311fc3c0d612807aad6bc72f4d64b54f)

## Author
Catalina G. Ivana — Operations & Process Optimization
