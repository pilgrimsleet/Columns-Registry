# Column Registry - Dynamic Spreadsheet Column Tracking

This is a more robust way to track columns that are referenced in scripts.  Hardcoding column positions, using named ranges, or searching for a hardcoded header string are all fragile.  
**Column Registry** maintains a central registry mapping that automatically detects and reconciles repositions and renames of columns so that you can reference the column position in scripts even if the names or position of the columns changes.

## Quick Start

```javascript
// 1. Register columns in ColumnRegistry sheet (manually, once)
// Script: myScript | Variable_ID: donor_name | Sheet: Donors | Header: Donor Name | Position: 1

// 2. Call at script start
function myScript() {
  updateColumns("myScript");
  const nameCol = getColumnPosition("myScript", "donor_name");
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    var name = data[i][nameCol - 1];  // Correct column, even if moved
    // ... use name
  }
}
```

## Installation

1. **Copy `columnRegistry.js`** into your Apps Script project
2. **Create a "ColumnRegistry" sheet** by running the performRegistryMaintenance() function from the AppScripts IDE.
3. **Register your columns** by adding rows to ColumnRegistry (e.g., Script=myScript, Variable_ID=donor_name, Sheet_Name=Donors, Header_Name="Donor Name", Position=1)
4. **Call `updateColumns("scriptId")` at the start** of any script that tracks columns

That's it. The registry handles the rest automatically.

## How It Works

### Three-Step Reconciliation Algorithm

When `updateColumns()` runs, it attempts to locate each registered column in three steps:

1. **Step 1: Three-way match** (name + position + tooltip)
   - If all three match, column is confirmed valid
   
2. **Step 2: Two-way match with confirmation** (name+position OR tooltip)
   - If found via name+position match or tooltip match, user confirms the location
   - In headless mode (no UI), automatically accepts the match
   
3. **Step 3: Fallback with suggestion** (name OR position only)
   - Offers best guess based on header name or last known position
   - Requires user confirmation or prompts for manual entry
   - Will return with an error if in headless mode

### Automatic Tooltip Tracking

Each registered column gets a tooltip in its header cell: `scriptId:variableId`

### Proactive Cross-Script Updates

When a tracked column relocates, the registry automatically updates all other scripts tracking the same column.

## API Reference

### Public Functions

#### `getColumnPosition(scriptId, variableId)`

Get a single column position from cache (no sheet access after first call).

```javascript
const col = getColumnPosition("myScript", "donor_name");
if (col > 0) {
  // Column exists and Last_Run is updated
  const value = row[col - 1];
}
```

**Returns:** 1-based column number, or -1 if not found

**Side effect:** Updates `Last_Run` timestamp for this entry

---

#### `getColumnPositions(scriptId)`

Get all columns for a script as an object (batch lookup).

```javascript
const cols = getColumnPositions("myScript");
// cols = { donor_name: 1, donation_amount: 5, ... }
```

**Returns:** Object mapping `variableId` → column number (or -1 if not found)

**Side effect:** Updates `Last_Run` timestamp for all entries with position > 0

---

#### `updateColumns(scriptId)`

Reconcile all column positions for a script. Handles 3-step matching and proactive updates.

```javascript
updateColumns("myScript");  // Call at script start
```

**Does NOT update `Last_Run`** — only getter functions do

**Side effects:**
- May prompt user to confirm/locate columns (in UI mode)
- Updates registry with new positions/names if columns moved
- Adds/updates tooltips in header cells
- Auto-updates other scripts tracking the same columns

---

#### `performRegistryMaintenance()`

Full maintenance: prune old entries, update all scripts, rebuild tooltips.

**Three phases:**

1. **Prune** – Removes entries not accessed for >13 months
2. **Update** – Runs `updateColumns()` for each registered script
3. **Rebuild** – Clears and rebuilds all tooltips from registry

```javascript
performRegistryMaintenance();  // Manual call or scheduled trigger
```

**Prompts user** for confirmation before pruning and on completion.

---

## Usage Examples

### Example 1: Basic Column Lookup

```javascript
function processDonations() {
  const SCRIPT_ID = "donorProcessor";
  updateColumns(SCRIPT_ID);  // Reconcile columns
  
  const donorCol = getColumnPosition(SCRIPT_ID, "donor_name");
  const amountCol = getColumnPosition(SCRIPT_ID, "amount");
  
  if (donorCol < 0 || amountCol < 0) {
    Logger.log("Required columns not found in registry");
    return;
  }
  
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const donor = data[i][donorCol - 1];
    const amount = data[i][amountCol - 1];
    // Process...
  }
}
```

### Example 2: Batch Lookup

```javascript
function updateDonorStats() {
  updateColumns("donorStats");
  const cols = getColumnPositions("donorStats");
  const values = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  
  for (let i = 1; i < values.length; i++) {
    const name = values[i][cols.name - 1];
    const email = values[i][cols.email - 1];
    // Process...
  }
}
```

## Last_Run Tracking

The `Last_Run` column tracks when each entry was last **accessed**.  Entries older than **13 months** are candidates for pruning by running performRegistryMaitenance().  13 months was chosen by default since the project that prompted this script is an accounting spreadsheet that has some scripts that are run on a yearly basis. You can change the PRUNE_THRESHOLD_MS constant to suit your use case.

## License

MIT

