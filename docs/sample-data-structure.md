# Data Structure — _XMLDATA Staging Sheet

## Overview
The `_XMLDATA` sheet is the output staging area populated by Power Query
after all 5 transformation queries have run. It feeds directly into the
VBA XML generator (`xml_generator.bas`).

All column references in the VBA code map to this structure.

---

## Column Map

### Client & Document Identity
| Column | Field Name | Type | Description |
|--------|-----------|------|-------------|
| L | ExitDate | Date | Shipment exit date — used for date filter and document grouping |
| O | ClientCode | Text | ERP client code — part of document grouping key |
| P | ClientName | Text | ERP client name — written to XML CustomerName |

### Transport & Document Fields
| Column | Field Name | Type | Description |
|--------|-----------|------|-------------|
| BC | TransportReason | Text | Reason for transport (e.g. "GOODS RETURNED FOR PROCESSING") |
| BD | GoodsAppearance | Text | Goods description (e.g. "FRESH MEAT") |
| BE | TransportInCharge | Text | Transport responsibility — mapped to "Sender" or "Recipient" |

### Product Lines
| Column | Field Name | Type | Description |
|--------|-----------|------|-------------|
| AX | ProcessingLineDesc | Text | Processing line description (conditional — written only if populated) |
| AY | ProcessingLineQty | Number | Processing line quantity |
| AZ | OfalLineDesc | Text | Offal line description (conditional — written only if populated) |
| BA | OfalLineQty | Number | Offal line quantity |
| BB | AggregatedData | Text | Main aggregated product description row |

### Traceability Fields (Dynamic String)
These columns are concatenated dynamically into a single traceability row.
Header row values are used as field labels in the output string.

| Column | Field Name | Description |
|--------|-----------|-------------|
| AV | AnimalTag | Animal identification tag |
| AW | Breed | Animal breed |
| AM | SlaughterNumber_Lot | Slaughter number / lot composite |
| AD | WeightKg | Carcass weight (kg) |
| AC | Category | Animal category |
| AN | Classification | Carcass classification |
| AR | Form4 | Regulatory form reference number |

### Notes
| Column | Field Name | Type | Description |
|--------|-----------|------|-------------|
| BF | Notes | Text | Free-text notes — written as NOTE row (conditional) |

---

## Document Grouping Logic

One XML `<Document>` block is generated per unique combination of:
- **ClientCode** (column O)
- **ExitDate** (column L, formatted as yyyymmdd)

Multiple rows sharing the same key produce multiple `<Row>` elements
within the same document block.

---

## Row Output Order (per document)
```
1. Processing line (AX/AY)     — conditional
2. Offal line (AZ/BA)          — conditional
3. Visual separator row
4. Aggregated data row (BB)    — conditional
5. Dynamic traceability string (AV, AW, AM, AD, AC, AN, AR)
6. Visual separator row        — conditional
7. Notes row (BF)              — conditional
```

---

## Data Flow
```
Excel Source Files
      ↓
01_clients_master_registry.pq
02_mapping_table.pq
03_traceability_register.pq
      ↓
04_extraction_and_enrichment.pq
      ↓
05_ddt_line_builder.pq
      ↓
_XMLDATA (staging sheet)
      ↓
xml_generator.bas
      ↓
ERP_DDT_Import.xml
```

---

## Notes for Reuse

- All string fields are sanitized via `CleanXML()` before XML output
- Decimal separators in quantity fields are normalized (comma → dot)
- Column references are positional — if columns are added/removed,
  update both this document and the VBA column map
- Date filter is applied at runtime via user input (InputBox)

> **OMS Backlog:** Replace positional column references with named ranges
> to make the macro resilient to structural changes in the staging sheet.
```

---

Struttura finale del repository completa:
```
etl-ddt-automation/
  README.md
  /power-query
    01_clients_master_registry.pq
    02_mapping_table.pq
    03_traceability_register.pq
    04_extraction_and_enrichment.pq
    05_ddt_line_builder.pq
  /vba
    xml_generator.bas
  /docs
    sample-data-structure.md  
