# Excel MCP Server Tools

This document provides detailed information about all available tools in the Excel MCP server.

## Workbook Operations

### create_workbook

Creates a new Excel workbook.

```python
create_workbook(filename: str) -> str
```

- `filename`: Path where to create workbook
- Returns: Success message with created file path

### create_worksheet

Creates a new worksheet in an existing workbook.

```python
create_worksheet(filename: str, sheet_name: str) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Name for the new worksheet
- Returns: Success message

### get_workbook_metadata

Get metadata about workbook including sheets and ranges.

```python
get_workbook_metadata(filename: str, include_ranges: bool = False) -> str
```

- `filename`: Path to Excel file
- `include_ranges`: Whether to include range information
- Returns: String representation of workbook metadata

## Data Operations

### write_data_to_excel

Write data to Excel worksheet.

```python
write_data_to_excel(
    filename: str,
    sheet_name: str,
    input_data: List[List],
    start_cell: str = "A1"
) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `input_data`: List of lists containing data to write
- `start_cell`: Starting cell (default: "A1")
- Returns: Success message

### read_data_from_excel

Read data from Excel worksheet.

```python
read_data_from_excel(
    filename: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str = None,
    preview_only: bool = False
) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Source worksheet name
- `start_cell`: Starting cell (default: "A1")
- `end_cell`: Optional ending cell
- `preview_only`: Whether to return only a preview
- Returns: String representation of data

## Formatting Operations

### format_range

Apply formatting to a range of cells.

```python
format_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int = None,
    font_color: str = None,
    bg_color: str = None,
    border_style: str = None,
    border_color: str = None,
    number_format: str = None,
    alignment: str = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Dict[str, Any] = None,
    conditional_format: Dict[str, Any] = None
) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Optional ending cell of range
- Various formatting options (see parameters)
- Returns: Success message

### merge_cells

Merge a range of cells.

```python
merge_cells(filename: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Success message

### unmerge_cells

Unmerge a previously merged range of cells.

```python
unmerge_cells(filename: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Success message

## Formula Operations

### apply_formula

Apply Excel formula to cell.

```python
apply_formula(filename: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to apply
- Returns: Success message

### validate_formula_syntax

Validate Excel formula syntax without applying it.

```python
validate_formula_syntax(filename: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to validate
- Returns: Validation result message

## Chart Operations

### create_chart

Create chart in worksheet.

```python
create_chart(
    filename: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data_range`: Range containing chart data
- `chart_type`: Type of chart (line, bar, pie, scatter, area)
- `target_cell`: Cell where to place chart
- `title`: Optional chart title
- `x_axis`: Optional X-axis label
- `y_axis`: Optional Y-axis label
- Returns: Success message

## Pivot Table Operations

### create_pivot_table

Create pivot table in worksheet.

```python
create_pivot_table(
    filename: str,
    sheet_name: str,
    data_range: str,
    target_cell: str,
    rows: List[str],
    values: List[str],
    columns: List[str] = None,
    agg_func: str = "mean"
) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data_range`: Range containing source data
- `target_cell`: Cell where to place pivot table
- `rows`: Fields for row labels
- `values`: Fields for values
- `columns`: Optional fields for column labels
- `agg_func`: Aggregation function (sum, count, average, max, min)
- Returns: Success message

## Worksheet Operations

### copy_worksheet

Copy worksheet within workbook.

```python
copy_worksheet(filename: str, source_sheet: str, target_sheet: str) -> str
```

- `filename`: Path to Excel file
- `source_sheet`: Name of sheet to copy
- `target_sheet`: Name for new sheet
- Returns: Success message

### delete_worksheet

Delete worksheet from workbook.

```python
delete_worksheet(filename: str, sheet_name: str) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Name of sheet to delete
- Returns: Success message

### rename_worksheet

Rename worksheet in workbook.

```python
rename_worksheet(filename: str, old_name: str, new_name: str) -> str
```

- `filename`: Path to Excel file
- `old_name`: Current sheet name
- `new_name`: New sheet name
- Returns: Success message

## Range Operations

### copy_range

Copy a range of cells to another location.

```python
copy_range(
    filename: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str = None
) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Source worksheet name
- `source_start`: Starting cell of source range
- `source_end`: Ending cell of source range
- `target_start`: Starting cell for paste
- `target_sheet`: Optional target worksheet name
- Returns: Success message

### delete_range

Delete a range of cells and shift remaining cells.

```python
delete_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up"
) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- `shift_direction`: Direction to shift cells ("up" or "left")
- Returns: Success message

### validate_excel_range

Validate if a range exists and is properly formatted.

```python
validate_excel_range(
    filename: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None
) -> str
```

- `filename`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Optional ending cell of range
- Returns: Validation result message

### create_download_link

Upload an Excel file to a Google Cloud Storage bucket and return a public download link.

```python
create_download_link(
    filename: str
) -> str
```

- `filename`: Path to the Excel file to upload
- Returns: Public URL to the uploaded file, or an error message if upload fails
