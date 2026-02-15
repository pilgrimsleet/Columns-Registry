/**
 * COLUMN REGISTRY - Centralized tracking of sheet column positions across scripts
 * Auto-reconciles when columns are moved, renamed, or deleted
 * 
 * Registry Sheet: "ColumnRegistry" with columns [Script, Variable_ID, Sheet_Name, Header_Name, Position, Last_Run, Notes]
 * Reconciliation: 1) Name+position → 2) Name+tooltip → 3) Position+tooltip → 4) User prompt
 * Proactive Updates: Column changes propagate to all tracking scripts automatically
 * 
 * Usage:
 *   Add entries manually to the ColumnRegistry sheet
 *   updateColumns("script");  // Call at script start
 *   const col = getColumnPosition("script", "var_id");  // Fast lookup
 */

/* Constants */

const REGISTRY_SHEET_NAME = "ColumnRegistry";
const PRUNE_THRESHOLD_MS = 13 * 30.44 * 24 * 60 * 60 * 1000; // 13 months in milliseconds

// Registry column indices (1-based)
const COL = {
  SCRIPT: 1,
  VARIABLE_ID: 2,
  SHEET_NAME: 3,
  HEADER_NAME: 4,
  POSITION: 5,
  LAST_RUN: 6,
  NOTES: 7
};

/* Public API */

/**
 * Get cached column position from registry (fast, no sheet access)
 * @param {string} scriptId - The script identifier
 * @param {string} variableId - The variable identifier
 * @returns {number} 1-based column position, or -1 if not found in registry
 */
function getColumnPosition(scriptId, variableId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const registrySheet = spreadsheet.getSheetByName(REGISTRY_SHEET_NAME);
  if (!registrySheet) throw new Error(`Registry sheet "${REGISTRY_SHEET_NAME}" does not exist`);
  const data = registrySheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const [script, varId] = data[i];
    if (script === scriptId && varId === variableId) {
      const position = Number(data[i][COL.POSITION - 1]) || -1;
      if (position > 0) {
        const row = i + 1;
        const timestampStr = new Date().toISOString();
        registrySheet.getRange(row, COL.LAST_RUN).setValue(timestampStr);
      }
      return position;
    }
  }
  
  Logger.log(`No registry entry for ${scriptId}.${variableId}`);
  return -1;
}

/**
 * Get all column positions for a script as an object
 * @param {string} scriptId - The script identifier
 * @returns {Object.<string, number>} Object mapping variable IDs to 1-based positions. Position is -1 if not found
 */
function getColumnPositions(scriptId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const registrySheet = spreadsheet.getSheetByName(REGISTRY_SHEET_NAME);
  if (!registrySheet) throw new Error(`Registry sheet "${REGISTRY_SHEET_NAME}" does not exist`);
  const data = registrySheet.getDataRange().getValues();
  const positions = {};
  const now = new Date().toISOString();
  
  for (let i = 1; i < data.length; i++) {
    const [script, varId] = data[i];
    if (script === scriptId) {
      const position = Number(data[i][COL.POSITION - 1]) || -1;
      positions[varId] = position;
      if (position > 0) {
        const row = i + 1;
        registrySheet.getRange(row, COL.LAST_RUN).setValue(now);
      }
    }
  }
  
  return positions;
}

/** Reconcile all positions for a script. Call at script start. Handles reconciliation and proactive updates */
function updateColumns(scriptId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const registrySheet = spreadsheet.getSheetByName(REGISTRY_SHEET_NAME);
  if (!registrySheet) throw new Error(`Registry sheet "${REGISTRY_SHEET_NAME}" does not exist`);
  const data = registrySheet.getDataRange().getValues();
  const now = new Date().getTime();
  
  for (let i = 1; i < data.length; i++) {
    const [script, varId, sheetName, headerName, position, lastRun, notes] = data[i];
    if (script !== scriptId) continue;
    
    const row = i + 1;
    const targetSheet = spreadsheet.getSheetByName(sheetName);
    if (!targetSheet) continue;
    
    const lastCol = targetSheet.getLastColumn();
    const currentHeaders = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const tooltips = targetSheet.getRange(1, 1, 1, lastCol).getNotes()[0];
    
    const registeredName = headerName || "";
    const oldPos = Number(position) || -1;
    const tooltipId = `${script}:${varId}`;
    const toColLetter = (idx) => String.fromCharCode(65 + idx);
    let newPosition = -1;
    let newName = registeredName;
    
    const matches = (idx) => ({
      name: registeredName && currentHeaders[idx] === registeredName,
      position: idx + 1 === oldPos,
      tooltip: (tooltips[idx] || "").includes(tooltipId)
    });
    
    const selectColumn = (idx) => {
      newPosition = idx + 1;
      newName = currentHeaders[idx];
      
      // Remove tooltip from old position if different
      if (oldPos > 0 && oldPos <= tooltips.length && oldPos !== newPosition) {
        const oldTooltip = tooltips[oldPos - 1] || "";
        if (oldTooltip.includes(tooltipId)) {
          const cleaned = oldTooltip.split(',').map(p => p.trim()).filter(p => p && p !== tooltipId).join(', ');
          targetSheet.getRange(1, oldPos).setNote(cleaned);
        }
      }
      
      // Add tooltip to new position if not present
      const existing = tooltips[idx] || "";
      if (!existing.includes(tooltipId)) {
        const newTooltip = existing ? existing.split(',').map(p => p.trim()).filter(p => p).concat(tooltipId).join(', ') : tooltipId;
        targetSheet.getRange(1, newPosition).setNote(newTooltip);
      }
    };
    
    // Step 1: All three match
    if (oldPos > 0 && oldPos <= currentHeaders.length) {
      const m = matches(oldPos - 1);
      if (m.name && m.position && m.tooltip) newPosition = oldPos;
    }
    
    // Step 2: (Name+Position) OR Tooltip - with confirmation
    if (newPosition <= 0) {
      let idx = null;
      let type = null;
      
      if (oldPos > 0 && oldPos <= currentHeaders.length) {
        const m = matches(oldPos - 1);
        if (m.name && m.position) { idx = oldPos - 1; type = "name+position"; }
      }
      
      if (idx === null) {
        for (let j = 0; j < tooltips.length; j++) {
          if (matches(j).tooltip) { idx = j; type = "tooltip"; break; }
        }
      }
      
      if (idx !== null) {
        const promptMsg = `Confirm column match for "${varId}":\n\nColumn ${toColLetter(idx)}: "${currentHeaders[idx]}"\nMatch type: ${type}\n\nEnter column letter to confirm (or change):`;
        try {
          const ui = SpreadsheetApp.getUi();
          const response = ui.prompt("Column Match Confirmation", promptMsg, toColLetter(idx), ui.ButtonSet.OK_CANCEL);
          if (response.getSelectedButton() === ui.Button.OK) {
            const userIdx = response.getResponseText().trim().toUpperCase().charCodeAt(0) - 65;
            if (userIdx >= 0 && userIdx < currentHeaders.length) selectColumn(userIdx);
          }
        } catch (e) {
          Logger.log(`Accepting ${type} match for ${varId} in headless mode`);
          selectColumn(idx);
        }
      }
    }
    
    // Step 3: Only name OR position - prompt or error
    if (newPosition <= 0) {
      let suggestion = null;
      for (let j = 0; j < currentHeaders.length; j++) {
        const m = matches(j);
        if ((m.name && !m.position) || (!m.name && m.position)) { suggestion = j; break; }
      }
      
      if (suggestion !== null || registeredName) {
        const defaultValue = suggestion !== null ? toColLetter(suggestion) : '';
        const suggestionText = suggestion !== null ? `\n\nSuggested: Column ${toColLetter(suggestion)}: "${currentHeaders[suggestion]}"` : '';
        const promptMsg = `Column "${registeredName || varId}" needs manual location.${suggestionText}\n\nEnter column letter (e.g., A, B, C):`;
        
        try {
          const ui = SpreadsheetApp.getUi();
          const response = ui.prompt(`Locate Column: ${varId}`, promptMsg, defaultValue, ui.ButtonSet.OK_CANCEL);
          if (response.getSelectedButton() === ui.Button.OK) {
            const idx = response.getResponseText().trim().toUpperCase().charCodeAt(0) - 65;
            if (idx >= 0 && idx < currentHeaders.length) selectColumn(idx);
          }
        } catch (e) {
          Logger.log(`Cannot resolve ${varId} in headless mode - returning error`);
        }
      }
    }
    
    if (newPosition <= 0) continue;
    
    registrySheet.getRange(row, COL.HEADER_NAME).setValue(newName);
    registrySheet.getRange(row, COL.POSITION).setValue(newPosition);
    
    if (oldPos === newPosition && registeredName === newName) continue;
    
    // Proactive update: when a column relocates, update all other scripts tracking the old position
    if (oldPos > 0 && oldPos !== newPosition) {
      for (let j = 1; j < data.length; j++) {
        const otherRow = j + 1;
        if (otherRow === row) continue;
        
        const [otherScript, otherVarId, otherSheet, otherHeaderName, otherPosition] = data[j];
        
        // If another script was tracking the same old column (by position and sheet)
        if (otherSheet === sheetName && Number(otherPosition) === oldPos) {
          registrySheet.getRange(otherRow, COL.POSITION).setValue(newPosition);
          registrySheet.getRange(otherRow, COL.HEADER_NAME).setValue(newName);
          registrySheet.getRange(otherRow, COL.NOTES).setValue(`Auto-updated: followed ${script}:${varId} from pos ${oldPos} to ${newPosition}`);
          Logger.log(`Proactively updated ${otherScript}.${otherVarId} from ${oldPos} to ${newPosition}`);
        }
      }
    }
  }
}

/* Registry Helpers */


/**
 * Full registry maintenance: prune unused entries, update all columns, rebuild tooltips
 * Runs in three phases:
 *   1. Prune entries not accessed for >13 months
 *   2. Update column positions for each script sequentially
 *   3. Rebuild all tooltips from cleaned registry
 */
function performRegistryMaintenance() {
  Logger.log("=== Starting Registry Maintenance ===");
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create registry sheet if it doesn't exist
  let registrySheet = spreadsheet.getSheetByName(REGISTRY_SHEET_NAME);
  if (!registrySheet) {
    registrySheet = spreadsheet.insertSheet(REGISTRY_SHEET_NAME);
    registrySheet.appendRow(["Script", "Variable_ID", "Sheet_Name", "Header_Name", "Position", "Last_Run", "Notes"]);
    registrySheet.getRange(1, 1, 1, COL.NOTES).setFontWeight("bold");
    registrySheet.setFrozenRows(1);
  }
  
  let data = registrySheet.getDataRange().getValues();
  
  const now = new Date().getTime();
  
  // Phase 1: Prune unused entries
  Logger.log("\nPhase 1: Pruning unused entries (>13 months)");
  const entriesToPrune = [];
  
  for (let i = 1; i < data.length; i++) {
    const [script, varId, sheetName, headerName, position, lastRun] = data[i];
    if (!lastRun || !script || !varId) continue;
    
    try {
      const ageMs = now - new Date(lastRun).getTime();
      if (ageMs > PRUNE_THRESHOLD_MS) {
        const ageMonths = (ageMs / (30.44 * 24 * 60 * 60 * 1000)).toFixed(1);
        entriesToPrune.push([i + 1, `${script}:${varId}`, ageMonths, sheetName]);
      }
    } catch (e) {
      Logger.log(`  Skipping row ${i + 1}: invalid date format`);
    }
  }
  
  let prunedCount = 0;
  if (entriesToPrune.length > 0) {
    const pruneList = entriesToPrune
      .map(([, entry, age, sheet]) => `• ${entry} (${age}mo, ${sheet})`)
      .join('\n');
    
    const ui = SpreadsheetApp.getUi();
    if (ui.alert(`Delete ${entriesToPrune.length} unused entry(ies)?\n\n${pruneList}`, ui.ButtonSet.YES_NO) === ui.Button.YES) {
      entriesToPrune.reverse().forEach(([rowNum]) => registrySheet.deleteRow(rowNum));
      prunedCount = entriesToPrune.length;
      Logger.log(`  Pruned ${prunedCount} entries`);
      data = registrySheet.getDataRange().getValues();
    } else {
      Logger.log("  Pruning cancelled by user");
      SpreadsheetApp.getUi().alert("✓ Maintenance cancelled");
      return;
    }
  } else {
    Logger.log(`  No entries to prune`);
  }
  
  // Phase 2: Update all column positions by script
  Logger.log("\nPhase 2: Updating column positions for each script");
  const scripts = new Set();
  for (let i = 1; i < data.length; i++) {
    const script = data[i][0];
    if (script) scripts.add(script);
  }
  
  for (const scriptId of scripts) {
    Logger.log(`  Updating: ${scriptId}`);
    updateColumns(scriptId);
  }
  
  // Phase 3: Rebuild all tooltips
  Logger.log("\nPhase 3: Rebuilding all tooltips");
  data = registrySheet.getDataRange().getValues(); // Refresh after column updates
  
  const sheetUpdates = {}; // Track updates by sheet: { sheetName: { colPos: ["script1:var1", "script2:var2"] } }
  
  for (let i = 1; i < data.length; i++) {
    const [script, varId, sheetName, headerName, position] = data[i];
    if (!sheetName || !position) continue;
    
    const pos = Number(position);
    if (pos <= 0) continue;
    
    if (!sheetUpdates[sheetName]) {
      sheetUpdates[sheetName] = {};
    }
    
    if (!sheetUpdates[sheetName][pos]) {
      sheetUpdates[sheetName][pos] = [];
    }
    
    sheetUpdates[sheetName][pos].push(`${script}:${varId}`);
  }
  
  let tooltipCount = 0;
  for (const sheetName in sheetUpdates) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`  ⚠ Sheet not found: ${sheetName}`);
      continue;
    }
    
    Logger.log(`  Processing sheet: ${sheetName}`);
    const lastCol = sheet.getLastColumn();
    
    // Clear all tooltips in header row first
    sheet.getRange(1, 1, 1, lastCol).clearNote();
    
    // Set new tooltips based on registry
    for (const pos in sheetUpdates[sheetName]) {
      const colPos = Number(pos);
      const scriptVarIds = sheetUpdates[sheetName][pos];
      const newTooltip = scriptVarIds.join(', ');
      sheet.getRange(1, colPos).setNote(newTooltip);
      tooltipCount++;
    }
  }
  
  const completionMsg = `✓ Maintenance complete:\n  • Pruned: ${prunedCount} entries\n  • Updated: ${scripts.size} script(s)\n  • Tooltips: ${tooltipCount} set`;
  Logger.log(`\n=== ${completionMsg.replace(/\n  • /g, " | ")} ===`);
  
  SpreadsheetApp.getUi().alert(completionMsg);
}
