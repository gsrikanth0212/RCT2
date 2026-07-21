# NEXT IMPLEMENTATION TASK — DATASOURCE RESOLUTION, RELATIONSHIP INTEGRITY & POWER BI MODEL VALIDATION

We are continuing an existing Tableau → Power BI migration project.

DO NOT restart the project.

DO NOT rewrite the existing migration engine.

DO NOT replace the current architecture.

Read and follow these files first:

1. `master_migration_prompt.md`
2. `MIGRATION_ENGINE_PROGRESS.md`
3. `tableau_pbi_server.py`
4. `report_migration_tool.py`

Use the current implementation as the baseline and continue from where the previous work stopped.

The project is being developed in Claude Code connected to the GitHub repository.

The UI files are finalized.

DO NOT modify:

* `Index.html`
* `Landing.html`

Only make necessary changes to:

* `tableau_pbi_server.py`
* `report_migration_tool.py`

Preserve all existing APIs, routes, UI behavior, source/target folder behavior and existing migration functionality.

====================================================
CURRENT TEST WORKBOOKS
======================

The `source` folder contains the following important test cases:

1. `sales_dashboard.twbx`

   * Tableau packaged workbook.
   * Connected to an external datasource.
   * Migration completes, but opening the generated Power BI report produces a relationship/model error.

2. `sales_dashboard_new.twb`

   * Standalone Tableau workbook.
   * Not a packaged `.twbx`.
   * Must be handled generically without requiring packaged data files.

3. `sales_dashboard_no_conn.twbx`

   * Packaged Tableau workbook containing an extract.
   * No external live datasource connection is required.
   * This is an important test case for the embedded Hyper extraction path.

4. `Finance Dashboard_v2026.1.twbx`

   * Existing known-good regression fixture.
   * User observed that this packaged workbook with an extract migrates successfully without the current datasource errors.

Do NOT hardcode any of these workbook names into the production migration logic.

They are test fixtures only.

====================================================
CURRENT ERRORS TO INVESTIGATE
=============================

ERROR 1 — EXTERNAL/LIVE DATASOURCE PATH

The generated Power BI report/query is attempting to access:

`C:\path\to\sample_super_store.xlsx`

and fails with:

`File or Folder: Could not find a part of the path 'C:\path\to\sample_super_store.xlsx'.`

Affected logical tables include:

* `superstore_data`
* `People`
* `Returns`

The same issue appears in multiple generated queries.

Investigate the complete datasource path resolution pipeline.

The generated Power BI M query must NEVER contain a placeholder or non-existent path such as:

`C:\path\to\sample_super_store.xlsx`

unless that is the actual source file supplied by the user.

The migration engine must determine the datasource location from the Tableau workbook/package metadata.

For a `.twbx`:

1. Inspect the packaged archive.
2. Identify whether the referenced datasource file is physically embedded in the package.
3. If the source file is embedded, extract or reference it using a valid migration-time strategy.
4. If the datasource is external and not embedded:

   * Resolve the path relative to the Tableau workbook/package location where possible.
   * Resolve relative paths correctly.
   * Preserve the original datasource definition where appropriate.
   * Do not silently replace the path with a fake sample path.
5. If the original external datasource cannot be accessed:

   * Do not generate a broken Power BI query.
   * Generate a clear migration warning/audit entry explaining that the source file was unavailable.
   * If possible, preserve the table schema and semantic model so relationships and visuals can still be generated.
6. The generated Power BI report must not contain fake placeholder paths.

Implement a generic datasource path resolver.

Do not special-case Superstore or any workbook name.

====================================================
ERROR 2 — STANDALONE .TWB DATASOURCE SCHEMA MISMATCH
====================================================

For `sales_dashboard_new.twb`, the migration reports:

`The column 'Regional Manager' of the table wasn't found.`

This indicates that the generated semantic model/query/model metadata references a column that was not actually created in the corresponding Power BI table.

Investigate the full pipeline:

Tableau XML
→ datasource parsing
→ datasource schema
→ column extraction
→ table creation
→ column-name mapping
→ relationship generation
→ calculated fields
→ Power BI model
→ visual bindings

Identify why a field referenced by Tableau metadata is not present in the generated Power BI table.

Possible causes to investigate include:

* Tableau logical field vs physical column mismatch.
* Aliased column names.
* Datasource captions vs physical names.
* Calculated fields incorrectly treated as physical columns.
* Relationship metadata referencing logical table names.
* Columns available in Tableau metadata but missing from extracted data.
* Column name normalization mismatch.
* Table splitting logic.
* Dimension-table creation logic.
* `_col_name_map` inconsistencies.
* Case sensitivity.
* Spaces and special characters.
* Duplicate field names across tables.

The fix must be generic.

Every field referenced by:

* a relationship
* a calculated field
* a measure
* a visual
* a filter
* a hierarchy
* a parameter
* a datasource query

must resolve against the final Power BI semantic model.

If a field is a calculated field, it must not be incorrectly expected to exist as a physical Power Query column.

If a field is a physical column, it must exist in the corresponding Power BI table.

If Tableau uses a logical alias/caption, maintain a reliable mapping:

Tableau Internal Name
→ Tableau Caption
→ Physical Source Name
→ Power BI Table
→ Power BI Column

====================================================
ERROR 3 — EMBEDDED EXTRACT / HYPER DATATYPE AND CARDINALITY ISSUES
==================================================================

For `sales_dashboard_no_conn.twbx`, the packaged workbook contains an embedded extract.

The following errors occur:

1.

`DAX comparison operations do not support comparing values of type Number with values of type Text.`

2.

`Column 'Region' in Table 'People' contains a duplicate value 'North' and this is not allowed for columns on the one side of a many-to-one relationship or for columns that are used as the primary key of a table.`

3.

`Column 'Order ID' in Table 'Returns' contains a duplicate value 'OrderID A' and this is not allowed for columns on the one side of a many-to-one relationship or for columns that are used as the primary key of a table.`

These are critical semantic-model issues.

The migration engine must NOT assume that every Tableau relationship can automatically become a Power BI one-to-many relationship.

Before creating a Power BI relationship:

1. Identify the Tableau relationship/join definition.
2. Identify the source and target tables.
3. Identify the join columns.
4. Inspect the actual extracted data when available.
5. Determine whether the candidate key is unique.
6. Determine whether the relationship can safely be represented as:

   * one-to-many
   * many-to-one
   * one-to-one
   * many-to-many
7. Determine cross-filter direction where it can be safely inferred.
8. Validate the relationship against actual extracted data.

If a column contains duplicate values, it MUST NOT be placed on the "one" side of a one-to-many relationship.

For example:

`People[Region]`

contains:

`North`
`North`
`South`
...

Therefore, it cannot be the unique one-side key.

Similarly:

`Returns[Order ID]`

contains duplicate Order IDs and therefore cannot be assumed to be the unique one-side key.

The relationship engine must detect this instead of blindly assigning cardinality.

====================================================
RELATIONSHIP CARDINALITY STRATEGY
=================================

Implement a generic relationship inference and validation layer.

For every relationship:

1. Extract Tableau relationship metadata.
2. Extract join keys.
3. Locate both tables in the final Power BI model.
4. Locate both columns using the canonical column mapping.
5. Compare data types.
6. Check uniqueness of each candidate key when data is available.
7. Determine valid Power BI cardinality.

Rules:

* If left key unique and right key non-unique:
  one-to-many.

* If left key non-unique and right key unique:
  many-to-one.

* If both unique:
  one-to-one.

* If neither unique:
  many-to-many, if supported safely.

Do not create a one-to-many relationship simply because Tableau metadata contains a join.

The generated Power BI model must reflect the actual data characteristics where data is available.

If data is unavailable, use Tableau metadata and conservative inference, but clearly record the uncertainty in the audit report.

====================================================
DATATYPE COMPATIBILITY
======================

Before creating relationships or DAX comparisons, validate data types.

For every relationship key:

* Text ↔ Text
* Integer ↔ Integer
* Decimal ↔ Decimal
* Date ↔ Date
* DateTime ↔ DateTime

If Tableau represents a field as one type but the extracted data produces another type, normalize the Power BI schema before generating relationships or DAX.

Do not silently compare:

Number ↔ Text

Do not blindly use VALUE() or FORMAT() as a workaround.

First determine the correct semantic type from:

1. Tableau metadata.
2. Hyper schema.
3. Actual extracted values.

Then normalize both sides consistently.

For DAX calculated fields and relationship logic, ensure both operands have compatible data types.

====================================================
EMBEDDED HYPER EXTRACTION
=========================

The existing Hyper extraction implementation is already working for at least some extract-based workbooks.

Do not replace it.

Investigate why:

`Finance Dashboard_v2026.1.twbx`

works successfully while:

`sales_dashboard_no_conn.twbx`

produces datatype and cardinality errors.

Compare the two workbooks structurally.

Determine whether the difference is caused by:

* Hyper schema.
* Multiple tables.
* Relationships.
* Key uniqueness.
* Data type inference.
* Table splitting.
* Extract table naming.
* Logical vs physical table mapping.
* Relationship cardinality inference.

Use this comparison to improve the generic engine.

Do NOT hardcode either workbook.

====================================================
DATASOURCE SCENARIO MATRIX
==========================

The migration engine must support these scenarios:

A. `.twb` with external datasource reference.

B. `.twbx` with embedded Excel/CSV datasource.

C. `.twbx` with embedded Hyper extract.

D. `.twbx` with external/live datasource reference.

E. `.twb` where no source data is available during migration.

For each scenario:

* Parse Tableau metadata.
* Reconstruct datasource schema.
* Reconstruct tables.
* Reconstruct relationships and joins.
* Generate a valid Power BI semantic model.
* Generate visuals against that model.
* Do not generate fake paths.
* Do not generate dangling model references.

The migration process does NOT need to migrate or copy the actual external source data unless it is already embedded in the Tableau package.

The requirement is:

Datasource metadata + tables/schema + joins + relationships + semantic model + visuals

must be migrated correctly.

====================================================
POWER BI MODEL INTEGRITY VALIDATION
===================================

Before writing the final `.pbit`:

Run a complete semantic-model validation.

Verify:

1. Every table referenced by a relationship exists.

2. Every relationship source column exists.

3. Every relationship target column exists.

4. Every relationship uses compatible datatypes.

5. Every one-side relationship key is unique.

6. No duplicate relationship IDs exist.

7. No dangling relationship references exist.

8. No invalid GUID references exist.

9. Every measure references valid tables/columns.

10. Every calculated column references valid tables/columns.

11. Every visual references valid semantic model objects.

12. Every query references valid datasource objects.

13. Every query references valid file paths or embedded data.

14. No placeholder datasource paths remain.

15. No unresolved table or column references remain.

16. No Power BI Desktop repair dialog is expected from known model inconsistencies.

The migration must fail validation before producing the final report if critical model corruption is detected.

Prefer producing a clear validation error over generating a corrupted `.pbit`.

====================================================
REGRESSION TESTING
==================

After implementation, test all available workbooks, including:

* Existing 8 regression workbooks.
* `superstore_profits.twbx`
* `superstore_profits_extract.twbx`
* `sales_dashboard.twbx`
* `sales_dashboard_new.twb`
* `sales_dashboard_no_conn.twbx`
* `Finance Dashboard_v2026.1.twbx`

Do not stop after one successful workbook.

For every workbook:

1. Parse.
2. Extract datasource metadata.
3. Extract data when available.
4. Build semantic model.
5. Build relationships.
6. Validate relationship cardinality.
7. Validate datatypes.
8. Generate Power BI report.
9. Inspect generated model for dangling references.
10. Verify generated report opens without model/relationship errors.
11. Verify existing visual migration behavior has not regressed.

Compare before/after results for existing workbooks.

====================================================
AUDIT REPORT
============

Extend the existing migration audit output to clearly report:

Datasource Type:

* Embedded Hyper
* Embedded Excel/CSV
* External File
* Live Connection
* Unknown

Datasource Path:

* Resolved
* Embedded
* External and available
* External and unavailable

Tables:

* Created
* Missing

Columns:

* Created
* Missing
* Aliased

Relationships:

* Detected
* Created
* Skipped
* Invalid

Join Keys:

* Source Table
* Source Column
* Target Table
* Target Column

Cardinality:

* One-to-One
* One-to-Many
* Many-to-One
* Many-to-Many
* Unknown

Datatype Compatibility:

* Valid
* Normalized
* Incompatible

Validation:

* PASS
* WARNING
* ERROR

This audit must make it immediately clear why a relationship or datasource could not be migrated.

====================================================
CRITICAL SUCCESS CRITERIA
=========================

The immediate goal of this increment is NOT to add more chart types.

The immediate goal is:

1. Fix datasource path resolution.

2. Fix datasource schema/column mapping.

3. Complete and harden relationship migration.

4. Correctly reconstruct simple Tableau data models in Power BI.

5. Prevent invalid relationship cardinality.

6. Prevent Number-vs-Text DAX errors caused by inconsistent schema typing.

7. Prevent dangling relationships and missing table/column references.

8. Ensure generated Power BI reports open successfully without relationship/model repair errors.

9. Preserve all existing visualization improvements and regression behavior.

10. Keep the solution generic and metadata-driven.

Do not add workbook-specific conditions.

Do not hardcode:

* workbook names
* sheet names
* table names
* field names
* datasource paths

====================================================
IMPLEMENTATION ORDER
====================

Implement in this exact order:

STEP 1
Read and understand the current relationship implementation, especially Increment 22.

STEP 2
Trace the full datasource lifecycle for all three scenarios:

TWB
TWBX + external/live datasource
TWBX + embedded Hyper extract

STEP 3
Fix datasource path resolution and remove fake placeholder paths.

STEP 4
Fix canonical table/column mapping and ensure every semantic-model reference resolves.

STEP 5
Fix datatype inference and normalization.

STEP 6
Implement relationship cardinality detection using actual extracted data when available.

STEP 7
Validate all relationships before generating the final PBIT.

STEP 8
Fix the specific "Relationship object doesn't exist" class of errors generically.

STEP 9
Run the complete regression suite.

STEP 10
Only after all reports pass semantic-model validation, update `MIGRATION_ENGINE_PROGRESS.md` with:

* Root causes found.
* Existing increments reused.
* New fixes implemented.
* Regression results.
* Remaining known limitations.
* Exact next priority.

Do not proceed to new visualization increments until this semantic-model and relationship validation phase is stable.

====================================================
IMPORTANT
=========

The fact that `Finance Dashboard_v2026.1.twbx` already migrates successfully while the new Superstore scenarios fail is an important regression comparison.

Use it to identify the actual structural difference between successful and failing datasource paths.

Do not assume the existing Hyper extraction is broken.

Do not assume relationships are completely missing.

The current code already contains relationship-related implementation and Increment 22 specifically added dimension tables and relationship objects.

The task is to determine what remains incomplete or incorrect in that implementation and harden it generically.

The final outcome of this increment must be:

Tableau Workbook
→ Datasource Detection
→ Datasource Resolution
→ Schema Extraction
→ Table Creation
→ Canonical Column Mapping
→ Datatype Validation
→ Relationship/Join Detection
→ Cardinality Inference
→ Power BI Semantic Model
→ Model Validation
→ Visual Generation
→ Valid `.pbit`

The Power BI report must be structurally valid before visual fidelity work continues.

```
```
