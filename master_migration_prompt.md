# MASTER ENGINEERING SPECIFICATION

## Universal Tableau (.twb/.twbx) → Power BI (.pbit/.pbix) Migration Engine

## 1. Project Continuity

You are continuing work on an existing Tableau → Power BI migration engine.

The project already contains:

* Python backend
* Migration engine
* Existing migration progress
* Tableau workbooks in the Source folder
* Generated Power BI reports in the Target folder
* Error screenshots in the Errors folder

Read and understand the current implementation before making changes.

Continue from the existing architecture.

Do NOT restart or rewrite the application.

Reuse existing functionality wherever possible.

Only refactor where necessary to improve architecture, correctness, maintainability, and migration quality.

---

## 2. Files That Must Not Be Modified

Do NOT modify:

* Index.html
* Landing.html

The UI is finalized.

All improvements must be implemented in the Python backend.

---

## 3. Primary Objective

Build a production-quality Tableau → Power BI migration engine capable of migrating **arbitrary Tableau workbooks**.

The engine must not be optimized for:

* Global Payments
* AB_NYC
* HR Analytics
* Sales Param

Those workbooks are only regression tests.

Tomorrow, I may place hundreds of completely different Tableau workbooks into the Source folder.

The migration engine must work without requiring workbook-specific code.

Avoid:

* hardcoded workbook names
* hardcoded sheet names
* hardcoded field names
* hardcoded chart mappings
* workbook-specific conditions

Everything must be dynamically inferred.

---

# 4. Migration Pipeline

Implement and validate the following pipeline.

1. Read Tableau Workbook

↓

2. Extract Datasources

↓

3. Extract Data Model

↓

4. Build Semantic Workbook Model

↓

5. Convert Calculated Fields

↓

6. Classify Visualization

↓

7. Generate Power BI Visual Specification

↓

8. Generate Power BI Report

↓

9. Validate Report

↓

10. Regression Testing

Every stage should be modular and independently testable.

---

# 5. Datasource Migration

Completely migrate Tableau datasource metadata.

Preserve:

* datasource names
* tables
* custom SQL
* joins
* relationships
* unions
* aliases
* data types
* default aggregation
* null handling
* geographic roles
* formatting
* extracts (where possible)

Do not lose datasource information.

---

# 6. Data Model Migration

The Power BI data model should closely resemble Tableau.

Preserve:

* relationships
* joins
* cardinality
* primary keys
* foreign keys
* calculated columns
* hierarchies
* sort columns
* date tables
* date hierarchies
* numeric precision
* currency
* percentage
* date
* datetime
* time

Fix the observed issue where Tableau dates (2016–2025) become 1900 in Power BI.

The original dates must always be preserved.

---

# 7. Semantic Workbook Model

Introduce a reusable Semantic Workbook Model that becomes the single source of truth.

Include:

* dimensions
* measures
* continuous fields
* discrete fields
* date fields
* categorical fields
* numeric fields
* calculated fields
* calculated field dependencies
* hierarchies
* groups
* sets
* bins
* parameters
* filters
* Measure Names
* Measure Values
* rows shelf
* columns shelf
* marks
* color
* size
* shape
* detail
* tooltip
* path
* analytics
* dashboard metadata
* layout metadata

All rendering components must consume this model instead of reading Tableau XML directly.

---

# 8. Calculated Fields

Convert Tableau calculated fields into reusable DAX.

Support:

* IF
* CASE
* IFNULL
* ISNULL
* ZN
* DATEADD
* DATEDIFF
* DATEPART
* DATENAME
* MAKEDATE
* LOOKUP
* INDEX
* RANK
* TOTAL
* RUNNING_SUM
* WINDOW_SUM
* WINDOW_AVG
* WINDOW_MIN
* WINDOW_MAX
* FIXED
* INCLUDE
* EXCLUDE

Support nested calculations and dependencies.

---

# 9. Visualization Intelligence

Do NOT determine chart type only from Tableau Marks.

Infer visualization using:

* rows
* columns
* Measure Names
* Measure Values
* dimensions
* measures
* color
* size
* shape
* detail
* tooltip
* filters
* analytics
* dashboard context
* continuous/discrete fields

Determine the analytical intent before generating the Power BI visual.

---

# 10. Supported Visualizations

Correctly identify and recreate:

* Table
* Matrix
* Card
* Multi-row Card
* KPI
* Bar
* Column
* Clustered Bar
* Clustered Column
* Stacked Bar
* Stacked Column
* 100% Stacked
* Line
* Multi-Line
* Area
* Stacked Area
* Combo
* Dual Axis
* Pie
* Doughnut
* Scatter
* Bubble
* Treemap
* Waterfall
* Ribbon
* Funnel
* Heat Map
* Highlight Table
* Gauge
* Filled Map
* Bubble Map
* Histogram
* Bullet
* Box Plot (best approximation)
* Small Multiples
* Calendar Heatmap
* Financial Tables
* Running Totals
* Reference Lines
* Trend Lines

If Power BI has no equivalent visual, generate the closest supported visualization.

Never fall back to a table unless unavoidable.

---

# 11. Automatic Field Mapping

Automatically populate Power BI field wells.

Examples:

Line Chart

* Axis
* Legend
* Values

Bar Chart

* Category
* Values
* Legend

Scatter

* X
* Y
* Size
* Details
* Legend

Treemap

* Category
* Values
* Group

Pie / Doughnut

* Legend
* Values
* Tooltips

KPI

* Correct aggregation

Do not generate empty templates requiring manual configuration.

---

# 12. Preserve Data Semantics

Correctly determine whether a field is:

* Dimension
* Measure
* Continuous
* Discrete
* Date
* Numeric
* Text
* Geographic
* Categorical

Use identical semantics in Power BI.

---

# 13. Dashboard Layout

Preserve:

* dashboard size
* floating objects
* containers
* nested containers
* spacing
* alignment
* padding
* titles
* text
* images
* legends
* filters

The generated Power BI dashboard should visually resemble Tableau.

---

# 14. Formatting

Preserve:

* colors
* fonts
* font sizes
* titles
* axis labels
* gridlines
* borders
* legends
* backgrounds
* conditional formatting
* number formatting
* currency
* percentages
* dates
* sorting

---

# 15. Resolve Known Issues

Address these issues generically:

* Line Charts incorrectly becoming Bar Charts
* Bar Charts incorrectly becoming Line Charts
* Multi-Line charts losing multiple series
* Doughnut charts missing legends and values
* Pie charts missing legends and values
* TreeMaps missing category/value assignments
* Scatter plots missing X, Y, Size, Details, Legend
* Multi-KPI cards requiring manual aggregation
* Date fields incorrectly becoming 1900
* Sorting not preserved
* Duplicate column names resolved incorrectly
* Incorrect identification of dimensions and measures

Fix these by improving the migration engine, not by adding workbook-specific exceptions.

---

# 16. Remove Duplicate Logic

There must be one shared implementation for:

* visualization classification
* category resolution
* measure resolution
* legend resolution
* aggregation
* date handling
* field mapping
* visual generation

Both worksheet rendering and dashboard rendering must use the same reusable components.

---

# 17. Windows Compatibility

Development is currently taking place on macOS.

Power BI Desktop is not fully supported on macOS.

However, the target runtime is Windows.

Requirements:

* Do not introduce macOS-specific behavior.
* Use cross-platform Python libraries.
* Use platform-independent file handling.
* Generate Power BI artifacts compatible with Windows Power BI Desktop.
* Validate package structure, report JSON, layouts, GUIDs, metadata, interactions, filters, bookmarks, and visual configuration.
* Reports should open in Windows Power BI Desktop without repair dialogs or manual fixes.

---

# 18. Regression Testing

Build an automated regression framework.

After every architectural change:

1. Convert every Tableau workbook in the Source folder.
2. Generate Power BI reports.
3. Validate every generated report.
4. Compare Tableau metadata with generated Power BI metadata.
5. Detect regressions before accepting changes.

Never improve one workbook while breaking another.

---

# 19. Migration Test Suite

Create and maintain `MIGRATION_TEST_SUITE.md`.

Track for every workbook:

* Workbook Name
* Worksheets
* Dashboards
* Datasources
* Relationships
* Calculated Fields
* Chart Inventory
* Generated Chart Types
* Data Model Status
* Layout Match
* Formatting Match
* Visual Similarity
* Unsupported Features
* Known Issues
* Resolved Issues
* Pending Improvements

This document should evolve as the migration engine improves.

---

# 20. Semantic Validation

Do not validate only the Power BI file structure.

Validate semantics.

For example:

If Tableau contains:

"Sales by Month"

The generated Power BI report must also represent:

"Sales by Month"

—not—

"Sales by Region"

Ensure dimensions, measures, aggregations, legends, sorting, and field assignments remain semantically correct.

---

# 21. Root Cause Analysis

Before implementing any fix:

Identify whether the issue originates from:

* datasource extraction
* data model
* metadata parsing
* semantic model
* calculated field conversion
* visualization classification
* field mapping
* DAX generation
* layout engine
* formatting engine
* Power BI JSON generation

Fix the root cause rather than symptoms.

---

# 22. Architecture First

When multiple bugs originate from duplicated code:

Refactor the architecture.

Do not duplicate fixes.

Prefer reusable components.

Reduce technical debt.

---

# 23. Code Quality

Maintain production-quality code.

* Modular architecture
* Clear separation of responsibilities
* Reusable components
* Minimal duplication
* Readable implementation
* Well-documented modules

Avoid large, deeply nested conditional blocks wherever practical.

---

# 24. Continuous Improvement Loop

Repeat until no major architectural improvements remain:

1. Convert all workbooks.
2. Validate reports.
3. Compare against Tableau metadata.
4. Identify differences.
5. Determine root cause.
6. Improve the migration engine.
7. Re-run regression tests.

---

# 25. Visual Comparison Validation

For every worksheet and dashboard:

* Compare generated Power BI visuals with the Tableau dashboard.
* Verify chart type, field assignments, legends, colors, layout, titles, filters, and sorting.
* If differences are found, trace them back to the migration engine and fix the underlying cause.

Do not rely solely on XML metadata—validate against the intended visual outcome.

---

# 26. Success Criteria

The final application should:

* Work for arbitrary Tableau workbooks.
* Require no workbook-specific code.
* Preserve datasource metadata.
* Recreate the data model.
* Convert calculated fields to DAX.
* Correctly distinguish dimensions and measures.
* Correctly classify chart types.
* Automatically populate Power BI field roles.
* Preserve layout, formatting, sorting, and dates.
* Generate valid Power BI reports.
* Open successfully in Windows Power BI Desktop.
* Achieve approximately 80% or higher visual and functional similarity for previously unseen Tableau dashboards.

## RELATIONSHIP, JOIN & DATA MODEL COMPLETION

```text
====================================================
RELATIONSHIP, JOIN & DATA MODEL COMPLETION
====================================================

This project is currently in the final implementation phase.

Do NOT restart the migration engine.

Continue from the latest implementation described in MIGRATION_ENGINE_PROGRESS.md and preserve every completed increment.

Before implementing any new feature, verify whether it has already been implemented partially. If it is partially implemented, complete and harden the existing implementation rather than replacing it.

----------------------------------------------------
RELATIONSHIP ENGINE COMPLETION
----------------------------------------------------

Confirm whether Tableau relationship migration has already been fully implemented.

If relationships are only parsed or partially generated, complete the implementation.

The migration engine must correctly reconstruct the Power BI semantic model including:

• Relationships
• Join Keys
• Join Columns
• Join Direction
• Cardinality
• Cross Filter Direction (when determinable)
• Primary / Foreign Key relationships
• Logical Layer
• Physical Layer

Relationship generation must be metadata-driven.

Never hardcode table names, workbook names or column names.

----------------------------------------------------
JOIN RECONSTRUCTION
----------------------------------------------------

Support generic reconstruction of Tableau joins.

Handle:

• Inner Join
• Left Join
• Right Join
• Full Outer Join
• Multi-table joins
• Multi-hop joins
• Logical relationships
• Physical joins
• Custom SQL references (when available)
• Multiple datasources

When Power BI does not have an equivalent implementation, generate the closest valid semantic model while preserving data correctness.

----------------------------------------------------
POWER BI DATA MODEL VALIDATION
----------------------------------------------------

The generated Power BI report must contain a valid semantic model.

Before completing migration verify:

• Every table exists.

• Every relationship references existing tables.

• Every relationship references existing columns.

• No dangling relationship objects exist.

• No orphaned tables exist.

• No duplicate relationship identifiers exist.

• No invalid relationship GUID references exist.

• Every visual references valid semantic model objects.

If any validation fails, automatically repair the semantic model before writing the final Power BI package.

The generated report must open successfully in Windows Power BI Desktop without repair dialogs or errors such as:

"Relationship object doesn't exist."

----------------------------------------------------
DATA MODEL VERIFICATION
----------------------------------------------------

For every migrated workbook generate a validation report confirming:

✓ Tables migrated

✓ Relationships migrated

✓ Join definitions migrated

✓ Calculated Fields migrated

✓ Measures generated

✓ DAX validated

✓ Hierarchies generated

✓ Parameters generated

✓ Semantic model validated

✓ Visuals successfully bound to the semantic model

----------------------------------------------------
GENERIC IMPLEMENTATION
----------------------------------------------------

Never implement fixes that only work for Sales Dashboard or any specific workbook.

The implementation must work for arbitrary future Tableau workbooks placed into the Source folder.

Every improvement must increase migration quality across all dashboards.

----------------------------------------------------
CURRENT PROJECT PRIORITY
----------------------------------------------------

The highest priority is no longer adding new chart types.

The highest priority is completing a reliable generic semantic model and relationship engine so that all generated visuals bind to a valid Power BI data model.

Once the semantic model is complete, continue improving visualization accuracy using that shared semantic model rather than workbook-specific logic.

This work should be completed by extending the current implementation, not by rewriting existing architecture.
```


## CONTINUATION, ARCHITECTURAL PRIORITY & QUALITY ASSURANCE

```
====================================================
CONTINUATION, ARCHITECTURAL PRIORITY & QUALITY ASSURANCE
====================================================

IMPORTANT

This project is already significantly implemented.

The existing migration engine, Python codebase and MIGRATION_ENGINE_PROGRESS.md represent the current baseline.

Before making any code changes:

1. Read the entire existing Python codebase.

2. Read MIGRATION_ENGINE_PROGRESS.md completely and treat it as the current implementation status.

3. Continue from the latest completed architecture.

4. Do NOT restart, redesign or replace already working modules unless there is a clear architectural benefit.

5. Preserve all previously implemented fixes and ensure backward compatibility with every Tableau workbook that currently migrates successfully.

6. Every improvement must extend the existing migration engine rather than introduce workbook-specific logic.

----------------------------------------------------
ARCHITECTURAL PRIORITY
----------------------------------------------------

When deciding between:

• fixing an individual workbook or chart

OR

• improving the shared migration architecture

Always prioritize improving the shared architecture.

The objective is to build a reusable commercial-grade Tableau → Power BI migration framework, not to optimize for individual workbooks.

----------------------------------------------------
SEMANTIC MODEL FIRST
----------------------------------------------------

The migration engine must become Semantic Model driven.

Visual renderers must NEVER read Tableau worksheet XML directly.

Instead, every renderer must consume a shared Semantic Workbook Model.

The Semantic Workbook Model must become the single source of truth for:

• Datasources
• Physical Tables
• Logical Tables
• Relationships
• Joins
• Dimensions
• Measures
• Data Types
• Continuous / Discrete fields
• Hierarchies
• Groups
• Sets
• Parameters
• Filters
• Calculated Fields
• Calculated Columns
• LOD Expressions
• Measure Names
• Measure Values
• Shelf assignments
• Dashboard metadata
• Layout metadata

Every chart should obtain Category, Values, Legend, Axis, Tooltip, Size, Detail and Aggregation information from this semantic model instead of directly interpreting Tableau XML.

----------------------------------------------------
COMPLETE DATASOURCE RECONSTRUCTION
----------------------------------------------------

Do NOT simply copy or reference the underlying source tables.

Reconstruct the complete Tableau datasource inside Power BI.

This includes:

• Physical Layer
• Logical Layer
• Relationships
• Joins
• Custom SQL
• Aliases
• Hidden Fields
• Data Types
• Default Aggregations
• Default Formatting
• Sort Orders
• Hierarchies
• Groups
• Sets
• Parameters
• Calculated Fields
• Calculated Columns
• Date Hierarchies

The Power BI semantic model must be generated from this reconstructed datasource before any visual generation begins.

----------------------------------------------------
VISUAL GENERATION
----------------------------------------------------

Visual generation should always follow this workflow:

Tableau Workbook

↓

Datasource Reconstruction

↓

Power BI Data Model

↓

Semantic Workbook Model

↓

Visualization Classification

↓

Field Role Resolution

↓

Power BI Visual Specification

↓

Power BI Report Generation

↓

Validation

No visual should be generated by directly reading Tableau worksheet metadata.

----------------------------------------------------
VISUAL VALIDATION
----------------------------------------------------

Do not validate only Power BI JSON or metadata.

Compare the generated Power BI dashboard against the Tableau dashboard.

Validate:

• Chart Type
• Field Assignments
• Axis
• Legend
• Values
• Colors
• Sorting
• Filters
• Layout
• Titles
• Formatting
• Dashboard Structure

Identify differences, determine the root cause and improve the migration engine rather than applying workbook-specific fixes.

----------------------------------------------------
REGRESSION TESTING
----------------------------------------------------

After every architectural change:

1. Automatically migrate every Tableau workbook available in the Source folder.

2. Generate Power BI reports.

3. Validate every generated report.

4. Ensure no regressions are introduced.

5. Update MIGRATION_TEST_SUITE.md with:

• Workbook Name
• Datasource Status
• Data Model Status
• Calculated Field Status
• Chart Detection
• Generated Visuals
• Layout Match
• Formatting Match
• Visual Similarity
• Overall Migration Score
• Known Issues
• Newly Fixed Issues
• Remaining Improvements

----------------------------------------------------
QUALITY TARGET
----------------------------------------------------

Every completed iteration should increase the overall migration quality for ALL Tableau workbooks, not just the workbook currently being tested.

The objective is to continuously improve the migration engine until it consistently achieves approximately 80% or higher visual and functional similarity for previously unseen Tableau dashboards while generating Power BI reports that open successfully in Windows Power BI Desktop without manual modifications or repair prompts.
```

Treat this as a commercial-grade migration product. Every architectural improvement should increase the quality of all future Tableau-to-Power BI migrations, not just the current sample workbooks.

