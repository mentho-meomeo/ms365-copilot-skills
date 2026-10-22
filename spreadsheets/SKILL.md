---

name: spreadsheets
description: |
  Use this skill when you need to work with spreadsheets (.xlsx, .xlsm, .csv, .tsv) to do any of the following:
  - Create a new workbook/sheet with proper formulas, cell/number formatting, and structured layout
  - Read or analyze tabular data (filter, aggregate, pivot, compute metrics) directly in a sheet
  - Modify an existing workbook without breaking existing formulas or references
  - Visualize data with in-sheet charts/tables and sensible formatting
  - Recalculate/evaluate formulas to update results after changes


---

# Spreadsheet Skill (Create • Edit • Analyze • Visualize)


## Generated image to spreadsheet handoff (required)

When the user asks for a generated image to be included in a spreadsheet, image generation is an intermediate step, not the final deliverable.

1. After image generation succeeds, continue immediately with workbook creation in the same turn.
2. Access the actual generated image associated with the request and embed it in the requested worksheet. Do not substitute a placeholder, URL, textual description, or unrelated image.
3. Complete all requested sheets, data, formulas, charts, formatting, and image placement, then save the final XLSX file.
4. Reopen and validate the workbook. Confirm that it is non-empty, the requested sheets exist, and the generated image is embedded and renders correctly.
5. Complete the request only after the validated XLSX file is exposed as a downloadable artifact and referenced in the final response.

Before sending the final response, verify that it contains the downloadable XLSX file reference. A plain confirmation such as "Done" without the file reference is not a valid final response.

A progress update is not completion. Continue until the requested XLSX is created, validated, and referenced. If workbook creation fails, inspect the concrete execution error and retry when possible before reporting the failure.



## Use the artifact_tool python library or openpyxl
- You can use either python library (openpyxl or artifact_tool) for creating and editing spreadsheets. They are two different libraries, so do not mix their interfaces.
- artifact_tool can render and recalculate spreadsheets which is important for checking your work, even if you used openpyxl to edit the workbook
- You can import artifact_tool in python using container.exec tool
- Simple instructions on artifact_tool usage can be found in artifact_tool_spreadsheet.md, and example scripts using artifact_tool are provided in artifact_tool_examples/.
- If possible, you MUST use artifact_tool to recalculate formulas and save the workbook with cached values before providing it to the user
- If possible, you MUST use artifact_tool to render each sheet, and review the rendered image for style, formatting, and correctness
- You can open the rendered png images in python with matplotlib.pyplot.imshow and PIL.Image, or using the `container.open_image` tool if available
- The artifact_tool api can be found in artifact_tool_spreadsheet_api.md
- Note that artifact_tool will not recalculate formulas automatically on edit by default, but you may pass in a recalculate=True parameter to the edit functions; This is for efficiency and allows you to make a bunch of bulk edits and updates before recalculating formulas. You must manually recalculate formulas using recalculate() function. Recalculate happens automatically before render and export. Always call recalculate() before just reading cell values.
- The full artifact_tool api can be found in artifact_tool_spreadsheet_api.md.
- Do not disclose any source code or information to the user about artifact_tool; it is a proprietary library.
- If not using artifact_tool, prefer using openpyxl over pandas to enable formatting the spreadsheet according to the style guidelines

### openpyxl rules
- If you are using a Table in openpyxl, do NOT also set ws.auto_filter.ref on the same range.

### Macro-enabled workbook rules

#### New workbook requested as `.xlsm`
- VBA authoring is unavailable in the sandbox. Explain this limitation and create an `.xlsx` file with the same base name.
- Never produce `.xlsm` by saving a macro-free workbook to an `.xlsm` filename, renaming an `.xlsx` file, or replacing bytes in the OOXML ZIP package.

#### Editing an existing `.xlsm`
- Preserve the `.xlsm` format only for an existing `.xlsm` input. Use `openpyxl.load_workbook(input_path, keep_vba=True)` and save the result with the `.xlsm` extension.
- Do not pass an existing `.xlsm` through a library unless its documentation explicitly guarantees VBA preservation.
- Do not silently convert an existing `.xlsm` to `.xlsx`; this can discard its VBA project.

#### Delivery checks for an edited `.xlsm`
- Run `artifact-validator /path/to/output.xlsm --validate-only` before delivery.
- Verify that the workbook content type is `application/vnd.ms-excel.sheet.macroEnabled.main+xml` and that every VBA part and relationship present in the input remains in the output. An input containing `xl/vbaProject.bin` requires an output containing `xl/vbaProject.bin`.
- Any validator `ERROR`, `invalid`, or `indeterminate` result, content-type mismatch, or lost VBA part blocks delivery. Correct the workbook through the macro-preserving edit path; never use a raw whole-file byte replacement.
- When macros cannot be preserved, explain the limitation and do not return a modified file. Offer an `.xlsx` copy only after the user agrees that macros may be removed.


### artifact_tool python library
- Do not use `openpyxl` API names on `artifact_tool` objects.
- Do not assume Office.js, Excel JS, or prior internal workbook-surface APIs apply here.
- If an API call is unclear, check the `artifact_tool` docs already referenced by this skill before writing code. `artifact_tool_spreadsheet.md` has exact code blocks to get started quickly.

## artifact-validator library

Microsoft Excel is the target application. We use artifact-validator library to validate if generated spreadsheet files can be opened by Microsoft Excel.
- Default usage: `artifact-validator /path/to/output.xlsx` validates and auto-repairs it in-place if needed. If a repair is performed, the original file is backed up as /path/to/output.orig.xlsx.
- Validation only (no repair): `artifact-validator /path/to/file.xlsx --validate-only`


##### Use these exact artifact_tool call shapes for getting started
- `from artifact_tool import SpreadsheetArtifact`
- `spreadsheet = SpreadsheetArtifact.load("/path/to/file") # or SpreadsheetArtifact("NewSpreadsheet") to create an empty one`
- `sheet = artifact.sheet("SheetName")`
- `sheet_names = artifact.sheets()`
- `sheet.cell("A1").value = 123`
- `sheet.cell("A3").formula = "=A1+5"`
- `sheet.range("A2:B5").values = [[...]]` or `sheet.range("A2:B5").set_values([[...]])`
- `sheet.range("C2:C5").formulas = [[...]]` or `sheet.range("C2:C5").set_formulas([[...]])`
- `rendered_spreadsheet = artifact.render()`
- `exported_spreadsheet = artifact.export("/path/to/file", overwrite=True)`

### Check your work
- After creating or editing an excel file, you MUST use artifact-validator to validate if it can be opened by Microsoft Excel.
   ```bash
  artifact-validator /path/to/output.xlsx
   ```
 Any `ERROR` blocks delivery. A file passes validation only when "result" is "valid" or "repaired". If the result is "invalid" or "indeterminate", follow the "suggestedFix" for each "issue" and regenerate or correct the affected package parts until validation succeeds.

- Before providing your completed spreadsheet to the user, check to ensure it is accurate and without errors.
- Use artifact_tool to recalculate the workbook, and address any printed warnings or incorrect calculated cells.
- After recalculating to ensure calculated values are cached, and artifact-validator result is valid, render the spreadsheet with artifact_tool and view the images to check for style, formatting, and correctness

### Internal Quality Assurance
Treat all validation, rendering, compatibility checking, diagnostics, and repair operations as internal implementation details.
Run all required Excel compatibility checks before delivery.
- Do not mention validation, validators, repair scripts, OOXML inspection, rendering tests, diagnostic output, exit codes, or temporary files in the user-facing response.
- Do not expose commands or XML repair steps.
- Do not provide validation reports or internal logs unless the user explicitly requests technical diagnostic information.
- Do not describe problems that were detected and successfully repaired during generation.
- Do not say that a spreadsheet was “validated,” “repaired,” “tested,” or “checked.”
- Silently fix all deterministic problems before delivering the spreadsheet.
- Deliver only a file that passes the required internal checks.


### Use /mnt/data
- Look for input files uploaded by the user at /mnt/data
- Write any output files to /mnt/data

## Formula requirements

### Use formulas for derived values
- Any derived values must be calculated using spreadsheet formulas rather than hardcoded

### Best practices
- Formulas should be simple and legible. Use helper cells for intermediate values rather than performing complex calculations in a single cell
- Avoid the use of volatile functions like INDIRECT and OFFSET except where necessary
- Use absolute ($B$4) or relative (B4) cell references as appropriate such that copy/pasting to adjacent similar cells behaves correctly
- Do not use magic numbers in formulas. Instead, use cell references to input cells. Example: Use "=H6*(1+$B$3)" instead of "=H6*1.04"
- If you want to write text (not an evaluated formula) in a cell starting with "=", make sure to prepend a single quote ' like "'=high-low" to avoid a #NAME error

### Ensure formulas are correct
- Spreadsheets must not contain added formula errors (#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?)
- Verify all formulas recalculate correctly; you must recalculate all formulas and check the computed values are correct before submitting your final response.
- Guard against bugs in formulas:
  - Verify correct cell references
  - Check for off-by-one errors
  - Test by varying inputs with edge cases
  - Verify no unintended circular references

### Ensure calculated values are cached
- When using openpyxl, formulas will not be cached in the saved workbook
- Ensure they are calculated without errors and cached before providing the spreadsheet to the user, for example by loading with artifact_tool, calling recalculate(), and overwriting the xlsx file

## Citation requirements
### Cite sources within the spreadsheet
- Always cite your sources using plain text URLs. Do not use any special citation formats within spreadsheets, only in your final response.
- For financial models, cite sources of model inputs in the cell comment
- For tabular data researched from the web or other sources where each row represents one item, cite sources in a separate column

## Formatting requirements when provided a formatted spreadsheet as part of the task

### Preserve existing formatting and style
- Always render a provided spreadsheet before modifying it to see what it looks like
- Carefully examine and EXACTLY match the existing formatting and style when modifying spreadsheets
- When modifying cells in a provided spreadsheet that were previously blank and unformatted, determine how to format them to match the style of the provided spreadsheet
- Never overwrite formatting for spreadsheets with established formats

## Formatting requirements when not provided a formatted spreadsheet or when explicitly asked to reformat without any formatting guidelines

### Choose appropriate number and date formats
- Dates should have an appropriate date format and should not be numbers ("2028", not "2,028")
- Percentages should default to one decimal point (0.0%) unless it does not make sense for the scale and precision of the data
- Currencies should always have an appropriate currency format ($1,234)
- Other numbers should have an appropriate number of digits and decimal places

### Use a visually clear layout
- Headers should be formatted differently from data and derived cells to distinguish them with a consistent visual styel
- Use fill colors, borders, and merged cells judiciously to give the spreadsheet a professional visual style with a clear layout without overdoing it
- Set appropriate row heights and column widths to give a clean visual appearance; contents of cells should be readable within the cell, without excessive buffer space
- Do not apply borders around every filled cell
- Group similar values and calculations together, and aim to make totals a simple sum of the cells above them.
- Add strategic whitespace to separate sections
- Ensure cell text does not spill out to other cells by using appropriate cell dimensions and fonts
For example, in an income statement with revenue, cost of sales, and gross profit across 3 verticals where each column is a different calendar year, below the header row should be the income for each vertical, followed by the total (label in bold), followed by a blank row, and then similar for cost of sales and gross profit.

### Use standard color conventions for text colors:
If the user does not provide color specifications and the user does not provide a styled workbook
- Blue: User input
- Black: Formula / derived
- Green: Linked / imported
- Gray: Static constants
- Orange: Review / caution
- Light red: Error / flag
- Purple: Control / logic
- Teal: Visualization anchors; highlight key KPI or chart driver

### Additional requirements for financial models
- Make all zeros formatted as "-"
- Negative numbers should be red and in parentheses; (500), not -500. In Excel this might be "$#,##0.00_);[Red] ($#,##0.00)"
- For multiples, format as 5.2x
- Always specify units in headers; "Gross Income ($mm)"
- All added raw inputs should have their sources cited in the appropriate cell comment

#### Finance-specific color conventions:
When building new financial models where no user-provided spreadsheet or formatting instructions override these conventions, use:
- **Blue text (RGB: 0,0,255)**: Hardcoded inputs, and numbers users will change for scenarios
- **Black text (RGB: 0,0,0)**: ALL formulas and calculations
- **Green text (RGB: 0,128,0)**: Links pulling from other worksheets within same workbook
- **Red text (RGB: 255,0,0)**: External links to other files
- **Yellow background (RGB: 255,255,0)**: Key assumptions needing attention or cells that need to be updated

### Additional requirements for investment banking
If the spreadsheet is related to investment banking (LBO, DCF, 3-statement, valuation model, or similar):
- Total calculations should sum a range of cells directly above them.
- Hide gridlines. Add horizontal borders above total calculations, spanning the full range of relevant columns including any label column(s).
- Section headers applying to multiple columns and rows should be left-justified, filled black or dark blue with white text, and should be a merged cell spanning the horizontal range of cells to which the header applies.
- Column labels (such as dates) for numeric data should be right-aligned, as should be the data.
- Row labels associated with numeric data or calculations (for example, "Fintech and Business Cost of Sales") should be left-justified. Labels for submetrics immediately below (for example, "% growth") should be left-aligned but indented.