# 3ds Max Automation Plan

## Goal
Create an automated 3ds Max panel that:

- Accepts a `SERIES NAME` input (example: `B779`)
- Accepts furniture names in a text box (example: `B16`, `B1`, etc.)
- Loads an Excel package list and parses rows like `PKG026197 (B779B16,1;B779B1,1;B779-46,1;B779-92,1)`
- Automatically instantiates the selected furniture objects in the scene
- Groups or layers the arrangement under the package name, e.g. `PKG026197`
- Creates a layer named for the package and hides it with the layer eye icon

## Data model

A package row has this structure:
- `PKG026197` = package name
- `(B779B16,1;B779B1,1;B779-46,1;B779-92,1)` = object list
- `B779` = series name
- `B16`, `B1`, `46`, `92` = furniture identifiers inside the series

## Panel design

1. Series Name field
   - Single-line text box for the series code
2. Furniture names field
   - Multi-line text box or comma-separated input for furniture IDs
   - Example input: `B16,B1,46,92`
3. Package creation button
   - Reads the selected series and furniture values
   - Instantiates corresponding objects
4. Excel tab
   - File picker for the Excel workbook
   - Parse selected rows into package definitions
   - Display package rows in a list
5. Package action
   - Select a package row from the Excel list
   - Automatically create the arrangement in the scene
   - Build a layer named exactly like the package code
   - Hide the layer via the layer manager eye control

## Automation workflow

1. Read Excel file
   - Use Excel COM, `xlrd`, `openpyxl`, or 3ds Max Python Excel automation
   - Parse package rows into:
     - package code
     - series code
     - furniture IDs and quantities
2. Map furniture names to 3ds Max objects
   - Use a naming convention like `B779_B16`, `B779_B1`, `B779_46`, `B779_92`
   - Or keep a lookup table if actual object names differ
3. Instantiate objects
   - For each furniture item, create an instance or copy
   - Place instances in the scene layout
4. Create and configure layer
   - Make a layer named `PKG026197`
   - Assign created objects to that layer
   - Hide the layer by setting layer visibility off

## Example logic

- Input: `SERIES NAME = B779`
- Furniture input: `B16, B1`
- Create instances: `B779_B16`, `B779_B1`
- If Excel row selected: parse `PKG026197` and create all listed items
- Final output: layer `PKG026197` containing the arrangement and hidden automatically

## Suggested implementation

Use 3ds Max scripting with either:
- MAXScript for UI rollout and object creation
- Python in 3ds Max if preferred

### Recommended steps

1. Build a rollout UI with `dotNetControl` or `rollout` in MAXScript
2. Add:
   - `edittext` for series name
   - `edittext` for furniture IDs
   - `button` for package creation
   - `button` to load Excel file
   - `listbox` to show Excel package rows
3. On Excel selection, parse the package string
4. Instantiate objects and assign to a layer
5. Hide the layer programmatically

## Notes

- The package name should become the layer name.
- The series prefix is used to identify the correct furniture family.
- The Excel parser should support the exact format shown in the attachment.
- If quantity values are present, instantiate that many copies or use instances.

## Next step

Implement the panel in MAXScript or 3ds Max Python using the above UI and package parsing logic.
