# Migration Audit: sales_dashboard_no_conn

## Workbook Summary

| Metric | Value |
|--------|-------|
| Total Calculations | 0 |
| ✅ SUCCESS | 0 |
| ⚠️ PARTIAL | 0 |
| 🔄 FALLBACK | 0 |
| ❌ UNSUPPORTED | 0 |
| 💥 FAILED | 0 |
| Total Measures | 0 |
| LOD Calculations | 0 |
| Table Calculations | 0 |
| Potential Visual Issues | 0 |
| Potential Accuracy Risks | 0 |
| **Overall Confidence** | **100/100** |
| Failed: Empty Formula  | 0 |
| Failed: Translator     | 0 |
| Failed: Col Ref Fix    | 0 |
| Failed: Agg Fix        | 0 |
| Failed: Cyclic Dep     | 0 |
| Failed: Missing Dep    | 0 |
| Post-proc recovered    | 0 |

## Per-Worksheet Migration Quality Score

**Average Overall Match: 96/100** across 2 worksheet(s).

_Visual/Layout/Formatting % are heuristic estimates from known mapping quality and detected signals — not a pixel-level comparison against the original Tableau render. Calculation % and Filter % are measured directly from this workbook's translation results._

| Worksheet | Detected (Tableau) | Power BI Visual | Calc % | Visual % | Layout % | Format % | Filter % | Overall % |
|-----------|---------------------|------------------|--------|----------|----------|----------|----------|-----------|
| sales_by_year | Automatic | clusteredColumnChart | 100 | 95 | 100 | 80 | 100 | **96** |
| sales_by_region | Bar | clusteredColumnChart | 100 | 95 | 100 | 87 | 100 | **97** |

## Data Model

**3 table(s)**, **2 relationship(s)** in the generated Power BI semantic model. Every relationship below was validated against the actual generated tables/columns before the .pbit was written — none reference a table or column that doesn't exist.

### Tables

- superstore_data
- People
- Returns

### Relationships

| From | To | Cardinality | Cross-Filter | Active |
|------|-----|-------------|--------------|--------|
| superstore_data.Region | People.Region | many:one | oneDirection | True |
| superstore_data.Order ID | Returns.Order ID | many:one | oneDirection | True |

## Semantic Classification Breakdown

| Semantic Class | Count |
|----------------|-------|

## Full Calculation Detail

