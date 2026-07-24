# NEXT TASK — FINALIZE AGGREGATION SEMANTICS AND COMPLETE TABLEAU DATASOURCE / RELATIONSHIP MODEL MIGRATION

We are continuing the existing Tableau → Power BI migration project.

DO NOT restart the project.

DO NOT rewrite the migration engine.

DO NOT replace existing working implementations.

Read the following before making changes:

1. `master_migration_prompt.md`
2. `MIGRATION_ENGINE_PROGRESS.md`
3. `tableau_pbi_server.py`
4. `report_migration_tool.py`

Use the current implementation as the baseline.

The project is already significantly advanced. Continue from the latest completed increment and preserve all existing functionality.

Do NOT modify:

* `Index.html`
* `Landing.html`

Only modify the Python implementation where necessary.

The objective of this increment is to address two remaining semantic migration issues:

1. Tableau aggregation semantics are not being preserved correctly.
2. Tableau datasource tables and relationships are not always reconstructed as the equivalent Power BI data model.

====================================================
ISSUE 1 — TABLEAU AGGREGATION SEMANTICS
=======================================

Generic example:

In Tableau, suppose a user drags:

`Sales`

to the Rows shelf.

Tableau automatically interprets this as:

`SUM(Sales)`

The important semantic distinction is:

Physical field:

`Sales`

Aggregation:

`SUM`

The migration engine must preserve these as two separate concepts.

DO NOT treat:

`SUM(Sales)`

as a physical column named:

`Sum of Sales`

DO NOT treat:

`AVG(Sales)`

as a different physical source column.

The underlying source field remains:

`Sales`

while the aggregation is:

`SUM`

or:

`AVG`

or:

`MIN`

or:

`MAX`

or:

`COUNT`

or:

`COUNTD`

or:

`MEDIAN`

or another supported Tableau aggregation.

---

## REQUIRED SEMANTIC REPRESENTATION

For every Tableau measure used in a worksheet, preserve:

* Datasource
* Table
* Physical Column
* Tableau Field Name
* Tableau Caption
* Aggregation
* Calculation status
* Calculation formula, if applicable
* Shelf
* Axis role
* Visual role

Represent the concept internally as something equivalent to:

{
"table": "Orders",
"column": "Sales",
"aggregation": "SUM",
"role": "Rows"
}

The exact internal structure can follow the existing architecture.

Do NOT create a fake physical column called:

`Sum of Sales`

unless that is genuinely the source column name.

---

## TABLEAU DEFAULT AGGREGATION

When a numeric field is dragged to Rows or Columns in Tableau and no explicit aggregation override exists, preserve Tableau's actual aggregation semantics.

For example:

`Sales`

should become equivalent to:

`SUM(Sales)`

not:

`Sales`

as an unaggregated field.

The Power BI representation should preserve the equivalent behavior.

Where Power BI visual-level aggregation can faithfully represent the Tableau aggregation, use the appropriate aggregation metadata.

Where a dedicated Power BI measure is required, generate a measure equivalent to:

`SUM('Orders'[Sales])`

or:

`AVERAGE('Orders'[Sales])`

etc.

However, do not create duplicate measures unnecessarily.

The semantic model should distinguish:

Physical Column:

`Orders[Sales]`

from:

Measure:

`SUM('Orders'[Sales])`

and from:

Visual Aggregation:

`SUM`

Use the least intrusive representation that preserves Tableau behavior.

---

## AGGREGATION MAPPING

Verify and correctly map at least:

Tableau → Power BI

SUM → SUM

AVG → AVERAGE

MIN → MIN

MAX → MAX

COUNT → COUNT / COUNTROWS equivalent where appropriate

COUNTD → DISTINCTCOUNT

MEDIAN → MEDIAN

STDEV → STDEV.S / appropriate equivalent

STDEVP → STDEV.P / appropriate equivalent

VAR → VAR.S / appropriate equivalent

VARP → VAR.P / appropriate equivalent

ATTR → appropriate semantic equivalent or documented approximation

Do not blindly map functions if Tableau and Power BI semantics differ.

Use the existing DAX conversion and aggregation infrastructure wherever possible.

---

## AGGREGATION FIELD-NAME NORMALIZATION

Audit the entire migration pipeline for places where aggregation prefixes are being converted into field names.

Examples:

`sum:Sales`

`SUM(Sales)`

`Sales`

`Sum of Sales`

`avg:Profit`

`AVG(Profit)`

These may represent the same physical field with different semantic aggregation metadata.

Normalize them to a canonical representation:

Physical Field:

`Sales`

Aggregation:

`SUM`

Do not lose the aggregation while normalizing.

Do not accidentally create duplicate semantic fields:

`Sales`

and:

`Sum of Sales`

for the same underlying field.

This must also handle:

* `med`
* `median`
* `cnt`
* `count`
* `cntd`
* `countd`
* `avg`
* `sum`
* `min`
* `max`
* other supported Tableau aggregation prefixes

Use the existing `_clean_field_name`, aggregation resolution and measure registry logic where appropriate.

Do not create a parallel aggregation resolver if an existing shared resolver can be extended.

---

## CALCULATED FIELDS

If Tableau contains:

`SUM([Sales])`

inside a calculated field, do not incorrectly interpret `Sales` as an already-aggregated physical column.

Preserve the calculation semantics.

Distinguish between:

`[Sales]`

and:

`SUM([Sales])`

and:

`SUM([Sales]) / SUM([Profit])`

and other aggregate expressions.

The DAX generated must preserve the correct aggregation level as closely as Power BI permits.

---

## VISUAL VALIDATION FOR AGGREGATION

Create or use regression tests where the same physical field is used with different aggregations:

Sales → SUM

Sales → AVG

Sales → MIN

Sales → MAX

Sales → COUNT

Sales → COUNTD

Verify that:

1. The physical column remains `Sales`.

2. The aggregation is preserved separately.

3. The Power BI visual uses the correct aggregation.

4. The generated report does not create misleading fields such as `Sum of Sales` as if that were a physical source column.

5. Changing the Tableau aggregation results in the corresponding Power BI aggregation.

6. Existing visual migrations do not regress.

---

## ISSUE 2 — TABLEAU DATASOURCE TABLES AND RELATIONSHIPS

Generic Superstore example:

Tableau datasource:

`superstore_data`

contains physical tables:

`Orders`

`People`

`Returns`

The Tableau datasource may represent these tables using:

* Logical relationships
* Physical joins
* `<object-graph><relationships>`
* `<relation type="collection">`
* Nested `<relation type="table">`
* Other supported Tableau datasource XML structures

The migration engine must not treat:

`superstore_data`

as one physical Power BI table when Tableau actually contains multiple related physical tables.

The Power BI semantic model must preserve the actual table structure.

Expected result:

Power BI:

`Orders`

`People`

`Returns`

with the corresponding relationships recreated.

The outer Tableau datasource caption:

`superstore_data`

is metadata describing the datasource.

It must NOT replace the physical table names.

---

## TABLE EXTRACTION REQUIREMENT

For every Tableau datasource:

1. Identify the datasource.

2. Identify every physical table.

3. Identify every logical table.

4. Identify every relation.

5. Identify every join.

6. Identify every relationship.

7. Identify the source table for each relation endpoint.

8. Identify the target table.

9. Identify the source join key.

10. Identify the target join key.

11. Preserve the original table names and column mappings.

Do not collapse multiple physical tables into one Power BI table simply because they belong to the same Tableau datasource.

---

## SUPERSTORE EXPECTED MODEL

For the Superstore datasource, if Tableau metadata identifies:

Orders

People

Returns

then Power BI must contain these as separate tables.

The engine must not produce only:

superstore_data

People

Returns

while silently dropping Orders.

If `Orders` is the main/fact table, it must still be represented as an actual Power BI table.

The datasource caption:

`superstore_data`

may be retained as datasource metadata, but must not replace the physical table structure.

---

## RELATIONSHIP RECONSTRUCTION

For every Tableau relationship:

Extract:

* Source table
* Source column
* Target table
* Target column
* Join/relationship type
* Cardinality where available
* Relationship direction where available

Generate the corresponding Power BI relationship.

Do not guess table names.

Do not guess column names.

Do not hardcode:

Orders

People

Returns

or any other workbook-specific table names.

Use the actual Tableau XML metadata.

---

## RELATIONSHIP VALIDATION

Before generating the final Power BI model:

Verify:

1. Every Tableau physical table has a corresponding Power BI table.

2. Every relationship source table exists.

3. Every relationship target table exists.

4. Every relationship source column exists.

5. Every relationship target column exists.

6. Both relationship columns have compatible data types.

7. Cardinality is valid.

8. The "one" side is actually unique when data is available.

9. No relationship is silently dropped.

10. No table is silently merged into another table.

11. No table is silently omitted.

12. No dangling relationship objects exist.

13. No invalid relationship GUID references exist.

14. No visual references a field from a table that was not generated.

---

## MULTIPLE TABLE REPRESENTATIONS

The current migration engine has already discovered that different Tableau workbooks represent relationships differently.

Support both:

A.

`<object-graph><relationships>`

and:

B.

`<relation type="collection">`

with nested:

`<relation type="table">`

Do not assume that relationship migration only works when `<object-graph>` exists.

For older Tableau datasource structures, infer the physical table structure from the relation collection and the `<cols>` mappings.

Do not recursively collect every nested physical column into one flat logical table.

Only expose columns in a Power BI table if those columns are actually produced by that table's Power Query/M query or data extraction.

Maintain a canonical mapping:

Tableau Datasource

→ Tableau Logical Table

→ Tableau Physical Table

→ Power BI Table

→ Power BI Column

---

## HYPER EXTRACTS

For embedded Hyper extracts:

If one Hyper file contains multiple physical tables such as:

Orders

People

Returns

extract all available tables separately.

Do not stop after the first table.

Do not use synthetic fallback data when real Hyper data can be extracted.

Use the existing Hyper extraction implementation.

Preserve the improvements already made for multi-table Hyper extraction.

When actual data is available:

* Validate uniqueness.
* Validate datatypes.
* Validate relationship cardinality.

---

## EXTERNAL AND EMBEDDED DATASOURCES

This must work consistently for:

1. `.twb` with external datasource.

2. `.twbx` with embedded Excel.

3. `.twbx` with embedded Hyper.

4. `.twbx` with external datasource.

5. `.twb` without source data available.

The migration engine does NOT need to migrate external source data.

However, it MUST reconstruct:

* Datasource metadata
* Physical tables
* Logical tables
* Relationships
* Joins
* Columns
* Data types
* Semantic mappings

When actual data is available, use it to validate the model.

When actual data is unavailable, rely on Tableau metadata and clearly report uncertainty.

---

## CRITICAL ARCHITECTURAL REQUIREMENT

Do not solve these issues with workbook-specific fixes.

Do not add conditions such as:

if workbook == "Superstore"

or:

if table == "Orders"

or:

if field == "Sales"

The implementation must work for arbitrary Tableau workbooks.

The generic architecture must be:

Tableau Workbook

↓

Datasource Detection

↓

Datasource Structure Reconstruction

↓

Physical Table Detection

↓

Logical Table Detection

↓

Relationship / Join Detection

↓

Column Mapping

↓

Data Type Resolution

↓

Power BI Data Model

↓

Aggregation Semantics

↓

Calculated Fields / Measures

↓

Visual Field Binding

↓

Power BI Visual Generation

↓

Model Validation

---

## REGRESSION TESTS

Use the existing regression suite.

Add targeted regression coverage for:

1. A numeric Sales field dragged to Rows → SUM(Sales).

2. Same Sales field → AVG(Sales).

3. Same Sales field → MIN(Sales).

4. Same Sales field → MAX(Sales).

5. Count aggregation.

6. Distinct Count aggregation.

7. A datasource containing multiple related physical tables.

8. A relationship represented through `<object-graph>`.

9. A relationship represented through `<relation type="collection">`.

10. A multi-table Hyper extract.

11. A datasource where the outer datasource caption differs from physical table names.

For the Superstore-like model verify that:

Tableau:

Datasource = superstore_data

Tables = Orders, People, Returns

Relationships = Tableau-defined relationships

results in:

Power BI:

Tables = Orders, People, Returns

Relationships = equivalent Tableau relationships

No physical table should be missing.

---

## SUCCESS CRITERIA

This increment is complete only when:

✓ Tableau physical tables are preserved as Power BI tables.

✓ Tableau logical relationships are recreated as Power BI relationships where supported.

✓ Tableau physical joins are reconstructed appropriately.

✓ The datasource caption is not incorrectly used as a replacement for physical tables.

✓ Orders, People and Returns are all preserved in the Superstore test model when they exist in the Tableau datasource.

✓ No relationships reference missing tables.

✓ No relationships reference missing columns.

✓ No invalid relationship objects are generated.

✓ Aggregations remain separate from physical field names.

✓ `Sales` remains the underlying field.

✓ `SUM(Sales)` is represented as Sales + SUM aggregation.

✓ `AVG(Sales)` is represented as Sales + AVG aggregation.

✓ `COUNTD(Sales)` is represented as Sales + DISTINCTCOUNT semantics.

✓ Calculated fields preserve their aggregation semantics.

✓ Existing Hyper extraction improvements remain intact.

✓ Existing successful workbooks continue to migrate successfully.

✓ No workbook-specific hardcoding is introduced.

✓ All generated Power BI reports pass semantic-model validation.

✓ All existing chart and visual improvements remain intact.

====================================================
FINAL VALIDATION
================

After implementing the changes:

1. Run `py_compile`.

2. Run the full regression suite.

3. Run targeted aggregation tests.

4. Run targeted multi-table datasource tests.

5. Run Superstore datasource tests.

6. Inspect the generated Power BI model.

7. Verify all physical tables exist.

8. Verify all relationships exist.

9. Verify aggregation semantics.

10. Verify no dangling model references.

11. Verify no existing visual regressions.

Update `MIGRATION_ENGINE_PROGRESS.md` with:

* Root cause of Issue 1.
* Root cause of Issue 2.
* Existing implementation reused.
* New code changes.
* Tests executed.
* Regression results.
* Remaining known limitations.

Do not proceed with unrelated enhancements until these two semantic migration issues are fully validated.

The final architecture must treat:

PHYSICAL FIELD

and

AGGREGATION

as separate semantic concepts,

and must treat:

TABLEAU DATASOURCE

and

PHYSICAL TABLES

as separate concepts.

The datasource is a container/model definition.

The physical tables inside that datasource must be independently reconstructed in Power BI.

The objective is not merely to generate a visually similar Power BI report.

The objective is to preserve the Tableau workbook's underlying semantic meaning:

Datasource

→ Tables

→ Relationships

→ Joins

→ Columns

→ Aggregations

→ Calculations

→ Visuals.

```
```
