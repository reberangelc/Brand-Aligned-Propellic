Generate a branded Propellic spreadsheet for: $ARGUMENTS

You have access to the `propellic_sheets` MCP tool. Use it to generate and upload the spreadsheet.

## How to use the tool

1. Based on the user's request, plan the sheet structure
2. Call `propellic_sheets` with these parameters:
   - **title**: Workbook title
   - **sheets**: JSON array of sheet objects (see below)

## Sheet format

Each sheet in the JSON array should have:
- `name` (str): Tab name
- `type` (str): "summary", "data", or "reference"
  - Summary tabs appear first with Pink tab color
  - Data tabs get Midnight tab color
  - Reference/lookup tabs get Berry tab color
- `headers` (list[str]): Column header names
- `rows` (list[list]): Data rows
- `summary_rows` (list[int]): 0-based row indices to style as summary rows (Midnight background, white bold text)

## Spreadsheet design guidelines

- Always include a Summary tab first with high-level metrics
- Group related data into separate tabs
- Use clear, descriptive column headers
- Format numbers appropriately: dates as MM/DD/YYYY, currency as $#,##0, percentages as 0.0%
- Include totals/averages as summary rows where appropriate
- Keep column count manageable — split wide datasets across tabs

The tool returns a Google Drive link. Share it with the user.
