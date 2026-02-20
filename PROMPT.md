**MINIONS SHEETS — IMPLEMENTATION SPEC**

You are tasked with creating the complete initial foundation for `minions-sheets` — a structured spreadsheet system that brings modern Excel-like capabilities to the Minions ecosystem. This is "modern Excel as minions," providing structured tables with formulas, queries, and chart definitions.

---

**PROJECT OVERVIEW**

`minions-sheets` provides spreadsheet functionality through structured minion types. It allows developers and agents to create, query, and manipulate tabular data with formula evaluation, type-safe columns, SQL-like querying, and import/export capabilities for CSV and XLSX formats.

The core concept: spreadsheets should be structured, queryable, and agent-friendly. Agents can create spreadsheets to analyze data, run calculations, generate reports, and export results — all without needing a database. This is a massive unlock for data-driven agent workflows.

---

**CONCEPT OVERVIEW**

This project is built on the Minions SDK (`minions-sdk`), which provides the foundational primitives: Minion (structured object instance), Minion Type (schema), and Relation (typed link between minions).

A spreadsheet contains multiple sheets, each with typed columns and rows. Cells can contain values or formulas. Charts visualize sheet data. Cross-sheet references are supported via relations. The column type system reuses FieldType from minions-sdk for validation.

The system supports both TypeScript and Python SDKs with cross-language interoperability (both serialize to the same JSON format). All documentation includes dual-language code examples with tabbed interfaces.

---

**CORE PRIMITIVES**

This project defines the following Minion Types:

- `spreadsheet` — A container for multiple sheets with metadata and settings
- `sheet` — A single sheet (tab) with columns and rows
- `column` — A column definition with type, name, and validation rules
- `row` — A single row of data (can be represented as cells or as structured record)
- `cell` — An individual cell with value or formula
- `formula` — A formula definition (can be embedded in cell or standalone)
- `chart` — A chart visualization linked to sheet data

---

**MINIONS SDK REFERENCE — REQUIRED DEPENDENCY**

This project depends on `minions-sdk`, a published package that provides the foundational primitives. The GH Agent building this project MUST install it from the public registries and use the APIs documented below — do NOT reimplement minions primitives from scratch.

**Installation:**
```bash
# TypeScript (npm)
npm install minions-sdk
# or: pnpm add minions-sdk

# Python (PyPI) — package name is minions-sdk, but you import as "minions"
pip install minions-sdk
```

**TypeScript SDK — Core Imports:**
```typescript
import {
  // Core types
  type Minion, type MinionType, type Relation,
  type FieldDefinition, type FieldValidation, type FieldType,
  type CreateMinionInput, type UpdateMinionInput, type CreateRelationInput,
  type MinionStatus, type MinionPriority, type RelationType,
  type ExecutionResult, type Executable,
  type ValidationError, type ValidationResult,

  // Validation
  validateField, validateFields,

  // Built-in Schemas (10 MinionType instances — reuse where applicable)
  noteType, linkType, fileType, contactType,
  agentType, teamType, thoughtType, promptTemplateType, testCaseType, taskType,
  builtinTypes,

  // Registry — stores and retrieves MinionTypes by id or slug
  TypeRegistry,

  // Relations — in-memory directed graph with traversal utilities
  RelationGraph,

  // Lifecycle — CRUD operations with validation
  createMinion, updateMinion, softDelete, hardDelete, restoreMinion,

  // Evolution — migrate minions when schemas change (preserves removed fields in _legacy)
  migrateMinion,

  // Utilities
  generateId, now, SPEC_VERSION,
} from 'minions-sdk';
```

**Python SDK — Core Imports:**
```python
from minions import (
    # Types
    Minion, MinionType, Relation, FieldDefinition, FieldValidation,
    CreateMinionInput, UpdateMinionInput, CreateRelationInput,
    ExecutionResult, Executable, ValidationError, ValidationResult,
    # Validation
    validate_field, validate_fields,
    # Built-in Schemas (10 types)
    note_type, link_type, file_type, contact_type,
    agent_type, team_type, thought_type, prompt_template_type,
    test_case_type, task_type, builtin_types,
    # Registry
    TypeRegistry,
    # Relations
    RelationGraph,
    # Lifecycle
    create_minion, update_minion, soft_delete, hard_delete, restore_minion,
    # Evolution
    migrate_minion,
    # Utilities
    generate_id, now, SPEC_VERSION,
)
```

**Key Concepts:**
- A `MinionType` defines a schema (list of `FieldDefinition`s) — each field has `name`, `type`, `label`, `required`, `defaultValue`, `options`, `validation`
- A `Minion` is an instance with `id`, `title`, `minionTypeId`, `fields` (dict), `status`, `tags`, timestamps
- A `Relation` is a typed directional link (12 types: `parent_of`, `depends_on`, `implements`, `relates_to`, `inspired_by`, `triggers`, `references`, `blocks`, `alternative_to`, `part_of`, `follows`, `integration_link`)
- Field types: `string`, `number`, `boolean`, `date`, `select`, `multi-select`, `url`, `email`, `textarea`, `tags`, `json`, `array`
- `TypeRegistry` auto-loads 10 built-in types; register custom types with `registry.register(myType)`
- `createMinion(input, type)` validates fields against the schema and returns `{ minion, validation }` (TS) or `(minion, validation)` tuple (Python)
- Both SDKs serialize to identical camelCase JSON; Python provides `to_dict()` / `from_dict()` for conversion

**IMPORTANT:** Do NOT recreate these primitives. Import them from `minions-sdk` (npm) / `minions` (PyPI). Build your domain-specific types and utilities ON TOP of the SDK.

---

**WHAT YOU NEED TO CREATE**

**1. THE SPECIFICATION** (`/spec`)

Write a complete markdown specification document covering:

- Motivation and goals — why agents need structured spreadsheet capabilities
- Glossary of terms specific to spreadsheet operations
- Core type definitions for all seven minion types with full field schemas
- Column type system — reusing FieldType from minions-sdk (string, number, boolean, date, select, etc.)
- Formula syntax and evaluation engine — supported functions (SUM, AVG, COUNT, IF, VLOOKUP, etc.)
- Cell reference syntax — A1 notation, range notation (A1:B10), cross-sheet references
- Query language specification — SQL-like SELECT/WHERE syntax for sheet queries
- Import/export formats — CSV and XLSX specifications
- Chart types and configuration — line, bar, pie, scatter with data binding
- Cross-sheet reference semantics via `references` relations
- Best practices for agent-driven spreadsheet workflows
- Conformance checklist for implementations

**2. THE CORE LIBRARY** (`/packages/core`)

A framework-agnostic TypeScript library built on `minions-sdk`. Must include:
- **Unified Client Architecture**:
  - A standalone `MinionsSheets` client class that wraps all primitives and utilities in a unified facade.
  - A `SheetsPlugin` class that implements `MinionPlugin` for mounting onto the core `Minions` client (e.g. `minions.sheets`).
  - Both modular (direct imports) and centralized (client instance) usage must be supported.

- Full TypeScript type definitions for all sheet-specific types
- `SheetEngine` class — formula evaluation engine
  - `evaluate(formula, context)` — evaluate formula in given context
  - `parse(formula)` — parse formula into AST
  - Supported functions: SUM, AVG, COUNT, MIN, MAX, IF, AND, OR, NOT, VLOOKUP, HLOOKUP, INDEX, MATCH, CONCATENATE, LEN, TRIM, UPPER, LOWER, DATE, NOW, TODAY
  - Support for cell references (A1, B2) and ranges (A1:B10)
  - Support for cross-sheet references (Sheet2!A1)
  - Circular dependency detection
- `SheetQuery` class — SQL-like querying
  - `query(sheet, sql)` — execute SQL-like query on sheet data
  - Supports SELECT, WHERE, ORDER BY, LIMIT
  - Supports aggregation functions (SUM, AVG, COUNT, etc.)
  - Returns structured result set
- `SheetImporter` class — import from external formats
  - `fromCSV(csvString)` — parse CSV and create sheet structure
  - `fromXLSX(buffer)` — parse XLSX file and create spreadsheet structure
  - Auto-detect column types from data
  - Preserve formulas where possible
- `SheetExporter` class — export to external formats
  - `toCSV(sheetId)` — export to CSV string
  - `toXLSX(spreadsheetId)` — export to XLSX buffer
  - `toJSON(spreadsheetId)` — full structured JSON export
  - Preserve formulas and formatting
- `ColumnTypeValidator` — validate cell values against column types
  - Reuses FieldType validation from minions-sdk
  - Returns clear validation errors
- `ChartBuilder` class — create chart definitions
  - `createChart(type, dataRange, options)` — build chart configuration
  - Support for line, bar, pie, scatter chart types
  - Data binding to sheet ranges
- Clean public API with comprehensive JSDoc documentation
- Zero storage opinions — works with any backend

**3. THE PYTHON SDK** (`/packages/python`)

A complete Python port of the core library with identical functionality:
- **Unified Client Architecture**:
  - `MinionsSheets` standalone client class.
  - `SheetsPlugin` class for mounting onto the core `Minions` client.

- Python type hints for all classes and methods
- `SheetEngine`, `SheetQuery`, `SheetImporter`, `SheetExporter`, `ColumnTypeValidator`, `ChartBuilder` classes
- Same method signatures as TypeScript version (following Python naming conventions)
- Serializes to identical JSON format as TypeScript SDK (cross-language interoperability)
- Full docstrings compatible with Sphinx documentation generation
- Use `pandas` for data manipulation where appropriate
- Use `openpyxl` for XLSX import/export
- Published to PyPI as `minions-sheets`

**4. THE CLI** (`/packages/cli`)

A command-line tool called `sheets` that provides:

```bash
sheets new "Revenue Tracker"
# Interactively create a new spreadsheet

sheets import data.csv --sheet "Q1"
# Import CSV data into a new sheet

sheets query <id> "SELECT name, revenue WHERE revenue > 10000"
# Query sheet data with SQL-like syntax

sheets export <id> --format xlsx
# Export to XLSX file

sheets formula <cell-id> "=SUM(B2:B10)"
# Set formula for a specific cell

sheets add-column <sheet-id> --name "Total" --type number
# Add a new column to a sheet

sheets add-row <sheet-id> --data '{"name": "Acme", "revenue": 15000}'
# Add a new row with structured data

sheets chart <sheet-id> --type bar --range "A1:B10"
# Create a chart from sheet data

sheets validate <sheet-id>
# Validate all cells against column types

sheets cells <sheet-id> --range "A1:B10"
# Display cell range with values/formulas
```

Additional features:
- Interactive mode for creating spreadsheets with column definitions
- Table-formatted output for query results
- CSV/XLSX import with progress indicators
- Config file support (`.sheetsrc.json`) for default settings
- Colored output for validation errors

**5. THE DOCUMENTATION SITE** (`/apps/docs`)

Built with Astro Starlight. Must include:

- Landing page — "Modern Excel as minions" positioning, emphasize agent data workflows
- Getting started guide with both TypeScript and Python examples
- Core concepts:
  - Spreadsheet structure (spreadsheet → sheets → columns → rows → cells)
  - Column types and validation (FieldType system)
  - Formula syntax and evaluation
  - Query language (SQL-like SELECT/WHERE)
  - Cross-sheet references
  - Charts and visualizations
- API reference for both TypeScript and Python
  - Dual-language code tabs for all examples
  - Auto-generated from JSDoc/docstrings where possible
- Guides:
  - Creating and populating spreadsheets
  - Writing formulas (with function reference)
  - Querying sheet data
  - Importing from CSV and XLSX
  - Exporting to CSV and XLSX
  - Creating charts
  - Cross-sheet calculations
  - Agent-driven data analysis workflows
- CLI reference with example commands
- Formula function reference:
  - Math functions (SUM, AVG, COUNT, MIN, MAX, ROUND, etc.)
  - Logic functions (IF, AND, OR, NOT)
  - Lookup functions (VLOOKUP, HLOOKUP, INDEX, MATCH)
  - Text functions (CONCATENATE, LEN, TRIM, UPPER, LOWER, LEFT, RIGHT, MID)
  - Date functions (DATE, NOW, TODAY, YEAR, MONTH, DAY)
- Integration examples:
  - Using with data analysis agents
  - Building financial models
  - Creating reports from API data
  - Data transformation pipelines
- Best practices for spreadsheet design in agent workflows
- Contributing guide

**6. OPTIONAL: THE WEB APP** (`/apps/web`)

A visual spreadsheet playground (optional but recommended):

- Grid interface with Excel-like cell editing
- Formula bar with syntax highlighting
- Column type selector and validation indicators
- Live formula evaluation
- Query builder interface (visual SQL composer)
- Import/export interface for CSV and XLSX
- Chart editor with live preview
- Cross-sheet navigation
- Built with Next.js or SvelteKit with a grid library like AG Grid or Handsontable

---

**PROJECT STRUCTURE**

Standard Minions ecosystem monorepo structure:

```
minions-sheets/
  packages/
    core/                 # TypeScript core library
      src/
        types.ts          # Type definitions
        SheetEngine.ts
        SheetQuery.ts
        SheetImporter.ts
        SheetExporter.ts
        ColumnTypeValidator.ts
        ChartBuilder.ts
        formulas/         # Formula function implementations
          math.ts
          logic.ts
          lookup.ts
          text.ts
          date.ts
        index.ts          # Public API surface
      test/
      package.json
    python/               # Python SDK
      minions_sheets/
        __init__.py
        types.py
        sheet_engine.py
        sheet_query.py
        sheet_importer.py
        sheet_exporter.py
        column_type_validator.py
        chart_builder.py
        formulas/         # Formula implementations
          math.py
          logic.py
          lookup.py
          text.py
          date.py
      tests/
      pyproject.toml
    cli/                  # CLI tool
      src/
        commands/
          new.ts
          import.ts
          query.ts
          export.ts
          formula.ts
          add-column.ts
          add-row.ts
          chart.ts
          validate.ts
          cells.ts
        index.ts
      package.json
  apps/
    docs/                 # Astro Starlight documentation
      src/
        content/
          docs/
            index.md
            getting-started.md
            concepts/
            guides/
            formulas/     # Formula reference
            api/
              typescript/
              python/
            cli/
      astro.config.mjs
      package.json
    web/                  # Optional playground
      src/
      package.json
  spec/
    v0.1.md              # Full specification
  examples/
    typescript/
      simple-sheet.ts
      formula-evaluation.ts
      csv-import.ts
      query-example.ts
      chart-creation.ts
    python/
      simple_sheet.py
      formula_evaluation.py
      csv_import.py
      query_example.py
      chart_creation.py
  .github/
    workflows/
      ci.yml             # Lint, test, build for both TS and Python
      publish.yml        # Publish to npm and PyPI
  README.md
  LICENSE                # AGPL-3.0
  package.json           # Workspace root
```

---

**BEYOND STANDARD PATTERN**

These utilities and classes are specific to `@minions-sheets/sdk`:

**SheetEngine**
- Formula parser and evaluation engine
- Supports standard spreadsheet functions (SUM, AVG, IF, VLOOKUP, etc.)
- A1 notation for cell references (A1, B2, C3:D10)
- Cross-sheet references (Sheet2!A1)
- Circular dependency detection and error reporting
- Lazy evaluation for performance
- Formula caching for repeated calculations

**SheetQuery**
- SQL-like query language for sheets
- SELECT syntax: `SELECT column1, column2 WHERE condition ORDER BY column ASC LIMIT 10`
- Supports aggregations: `SELECT category, SUM(revenue) GROUP BY category`
- WHERE clause supports comparison operators (=, !=, <, >, <=, >=) and logic (AND, OR)
- Returns structured result set with metadata
- Query optimization for large datasets

**SheetImporter**
- CSV parsing with delimiter detection
- XLSX parsing using openpyxl (Python) or exceljs (TypeScript)
- Auto-detection of column types from data patterns
- Header row detection
- Formula preservation from XLSX
- Import validation with error reporting

**SheetExporter**
- CSV export with configurable delimiter
- XLSX export with formula preservation
- Style and formatting preservation (colors, fonts, borders)
- Multi-sheet export for spreadsheets
- JSON export with full structure including formulas and charts

**ColumnTypeValidator**
- Validates cell values against column FieldType
- Reuses validation logic from minions-sdk
- Returns clear error messages with cell references
- Bulk validation for entire columns or sheets
- Type coercion options (strict or lenient)

**ChartBuilder**
- Creates chart definition minions linked to sheets
- Supports chart types: line, bar, column, pie, scatter, area
- Data range binding (A1:B10 notation)
- Multiple series support
- Customizable styling (colors, labels, legends)
- Export to chart libraries (Chart.js, Recharts)

---

**CLI COMMANDS**

All commands with detailed specifications:

**`sheets new <title>`**
- Interactive spreadsheet creation wizard
- Asks for: spreadsheet name, initial sheet name, column definitions
- Creates `spreadsheet` minion with initial `sheet` minion
- Optionally creates `column` minions
- Returns created minion IDs

**`sheets import <file> --sheet <name>`**
- Imports CSV or XLSX file
- Auto-detects format from extension
- Creates sheet structure with auto-detected column types
- Shows import summary (rows, columns, validation errors)
- Accepts `--type-map` for explicit column type mapping

**`sheets query <id> <sql>`**
- Executes SQL-like query on sheet data
- Displays results in table format
- Accepts `--json` for JSON output
- Accepts `--limit` to limit result rows
- Returns exit code based on result count

**`sheets export <id> --format <csv|xlsx|json>`**
- Exports spreadsheet to specified format
- Outputs to file via `--output` flag or stdout
- Preserves formulas in XLSX export
- Includes all sheets in spreadsheet

**`sheets formula <cell-id> <formula>`**
- Sets formula for specific cell
- Validates formula syntax
- Shows evaluated result
- Accepts `--dry-run` to validate without saving

**`sheets add-column <sheet-id> --name <name> --type <type>`**
- Adds new column to sheet
- Type must be valid FieldType (string, number, boolean, date, etc.)
- Accepts `--required` and `--default` flags
- Returns created column ID

**`sheets add-row <sheet-id> --data <json>`**
- Adds new row to sheet
- Data as JSON object with column names as keys
- Validates against column types
- Returns created row ID

**`sheets chart <sheet-id> --type <type> --range <range>`**
- Creates chart from sheet data
- Type: line, bar, column, pie, scatter, area
- Range in A1 notation
- Accepts `--title` and `--style` flags
- Returns created chart ID

**`sheets validate <sheet-id>`**
- Validates all cells against column types
- Displays validation errors with cell references
- Returns error count and exit code
- Accepts `--fix` to attempt auto-correction

**`sheets cells <sheet-id> --range <range>`**
- Displays cell values and formulas for given range
- Format: table with cell references, values, and formulas
- Accepts `--values-only` to hide formulas
- Accepts `--formulas-only` to show only formula cells

---

**DUAL SDK REQUIREMENTS**

Critical cross-language compatibility requirements:

**Serialization Parity**
- Both TypeScript and Python SDKs must serialize minions to identical JSON format
- Field names, types, and structure must match exactly
- Relation types and metadata must be interchangeable
- Formula syntax must be identical across languages

**API Consistency**
- Same method names (adjusted for language conventions: TypeScript camelCase, Python snake_case)
- Same parameters and return types
- Same class hierarchies and interfaces
- Same formula function names and signatures

**Documentation Parity**
- Every code example in docs must have both TypeScript and Python versions
- Use Astro Starlight's code tabs: `<Tabs><TabItem label="TypeScript">...</TabItem><TabItem label="Python">...</TabItem></Tabs>`
- API reference must document both languages side by side

**Testing Parity**
- Shared test fixtures (JSON files with spreadsheet data) that both SDKs can consume
- Identical test case coverage for formula evaluation
- Cross-language integration tests (TypeScript SDK creates spreadsheet, Python SDK queries it)

---

**FIELD SCHEMAS**

Define these Minion Types with full JSON Schema definitions:

**`spreadsheet`**
```typescript
{
  id: string;
  title: string;
  description?: string;
  settings?: {
    locale?: string;           // For number/date formatting
    timezone?: string;
    defaultDateFormat?: string;
  };
  tags?: string[];
  createdAt: Date;
  updatedAt: Date;
}
```
Relations: `parent_of` → `sheet` minions

**`sheet`**
```typescript
{
  id: string;
  title: string;              // Sheet name (e.g., "Q1 Revenue")
  description?: string;
  rowCount?: number;          // Cached row count
  columnCount?: number;       // Cached column count
  frozen?: {                  // Frozen rows/columns
    rows?: number;
    columns?: number;
  };
  sortBy?: {                  // Default sort
    column: string;
    direction: 'asc' | 'desc';
  };
  createdAt: Date;
  updatedAt: Date;
}
```
Relations: `parent_of` → `column` minions
Relations: `parent_of` → `row` minions
Relations: `parent_of` → `chart` minions

**`column`**
```typescript
{
  id: string;
  title: string;              // Column name (e.g., "Revenue")
  description?: string;
  columnType: FieldType;      // Reuses FieldType from minions-sdk
  columnIndex: number;        // Zero-based index
  width?: number;             // Column width in pixels
  required?: boolean;
  defaultValue?: any;
  validation?: {              // Additional validation rules
    min?: number;
    max?: number;
    pattern?: string;
    options?: string[];       // For select types
  };
  createdAt: Date;
  updatedAt: Date;
}
```

**`row`**
```typescript
{
  id: string;
  title?: string;             // Optional row label
  rowIndex: number;           // Zero-based index
  data: Record<string, any>;  // Column ID/name → cell value
  metadata?: {
    hidden?: boolean;
    height?: number;
    style?: any;
  };
  createdAt: Date;
  updatedAt: Date;
}
```

**`cell`**
```typescript
{
  id: string;
  title?: string;             // Optional cell label (for named cells)
  cellReference: string;      // A1 notation (e.g., "B5")
  value?: any;                // Raw value
  formula?: string;           // Formula string (e.g., "=SUM(A1:A10)")
  evaluatedValue?: any;       // Cached formula result
  error?: string;             // Formula evaluation error
  format?: {                  // Cell formatting
    type?: 'number' | 'currency' | 'percentage' | 'date' | 'text';
    pattern?: string;
    color?: string;
    backgroundColor?: string;
    bold?: boolean;
    italic?: boolean;
  };
  createdAt: Date;
  updatedAt: Date;
}
```
Relations: `parent_of` → parent `row` or `sheet`
Relations: `references` → other `cell` minions (for formula dependencies)

**`formula`**
```typescript
{
  id: string;
  title: string;              // Formula name or description
  expression: string;         // Formula expression (e.g., "=SUM(A1:A10)")
  description?: string;
  parameters?: {              // Named parameters for reusable formulas
    name: string;
    type: string;
    defaultValue?: any;
  }[];
  createdAt: Date;
  updatedAt: Date;
}
```

**`chart`**
```typescript
{
  id: string;
  title: string;              // Chart title
  description?: string;
  chartType: 'line' | 'bar' | 'column' | 'pie' | 'scatter' | 'area';
  dataRange: string;          // A1 notation range (e.g., "A1:B10")
  series?: {                  // Multiple data series
    name: string;
    range: string;
    color?: string;
  }[];
  options?: {                 // Chart configuration
    xAxis?: { label?: string; };
    yAxis?: { label?: string; min?: number; max?: number; };
    legend?: { position?: 'top' | 'bottom' | 'left' | 'right'; };
    colors?: string[];
    stacked?: boolean;
  };
  createdAt: Date;
  updatedAt: Date;
}
```
Relations: `references` → `sheet` minion (data source)

---

**TONE AND POSITIONING**

This is a serious tool for data-driven agent workflows. Position it as:

- **Structured spreadsheets for agents** — not just storage, but queryable, formula-enabled data
- **Modern Excel as minions** — bring spreadsheet power to AI workflows
- **Zero database required** — agents can work with tabular data instantly
- **Production-ready** — built for real data analysis and reporting

Avoid:
- Trying to compete with Excel or Google Sheets on features
- Over-promising on performance for huge datasets
- Complexity for complexity's sake

The README should open with a concrete example: creating a sheet, adding data, running a query, and exporting results. Make it immediately tangible.

---

**INTEGRATION EXAMPLES**

Include working examples for:

**Data Analysis Agent** (TypeScript)
```typescript
import { SheetEngine, SheetQuery, SheetImporter } from '@minions-sheets/sdk';
import { Minion } from 'minions-sdk';

// Import CSV data
const importer = new SheetImporter();
const sheet = await importer.fromCSV(csvData);

// Query the data
const query = new SheetQuery();
const results = await query.query(sheet,
  'SELECT category, SUM(revenue) as total WHERE date > "2026-01-01" GROUP BY category'
);

// Apply formula
const engine = new SheetEngine();
const avgRevenue = await engine.evaluate('=AVG(B2:B100)', sheet);

console.log('Total by category:', results);
console.log('Average revenue:', avgRevenue);
```

**Financial Model** (Python)
```python
from minions_sheets import SheetEngine, SheetExporter
from minions_sdk import Minion

# Create sheet with formulas
spreadsheet = Minion.create('spreadsheet', {
    'title': 'Revenue Forecast'
})

sheet = Minion.create('sheet', {
    'title': 'Q1 Projection'
})

# Add columns and rows with formulas
engine = SheetEngine()
engine.set_cell(sheet.id, 'A1', 'Month')
engine.set_cell(sheet.id, 'B1', 'Revenue')
engine.set_cell(sheet.id, 'C1', 'Growth')
engine.set_cell(sheet.id, 'C2', '=B2/B1-1')  # Month-over-month growth

# Export to XLSX
exporter = SheetExporter()
buffer = exporter.to_xlsx(spreadsheet.id)
with open('forecast.xlsx', 'wb') as f:
    f.write(buffer)
```

---

**DELIVERABLES**

Produce all files necessary to bootstrap this project completely:

1. **Full specification** (`/spec/v0.1.md`) — complete enough to implement from
2. **TypeScript core library** (`/packages/core`) — fully functional, well-tested
3. **Python SDK** (`/packages/python`) — feature parity with TypeScript
4. **CLI tool** (`/packages/cli`) — all commands working with helpful output
5. **Documentation site** (`/apps/docs`) — complete with dual-language examples
6. **README** — compelling, clear, with concrete examples
7. **Examples** — working code in both TypeScript and Python
8. **CI/CD setup** — lint, test, and publish workflows for both languages

Every file should be production quality — not stubs, not placeholders. The spec should be complete. The core libraries should be fully functional. The docs should be ready to publish. The CLI should be ready to install and use.

---

**START SYSTEMATICALLY**

1. Write the specification first — nail down the field schemas, formula syntax, and query language
2. Implement TypeScript core library with SheetEngine, SheetQuery, and import/export
3. Port to Python maintaining exact serialization compatibility
4. Build CLI using the core library
5. Write documentation with dual-language examples throughout
6. Create working examples demonstrating key workflows
7. Write the README with concrete use cases

This is a foundational data tool for the Minions ecosystem. Agents will use this to analyze data, build models, and generate reports. Get it right.
