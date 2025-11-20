/**
 * Configuration for the read-only source sheet.
 * @typedef {Object} SourceConfig
 * @property {string} sheetName - The name of the source sheet.
 * @property {number} keyCol - The 1-based column index containing the unique ID.
 * @property {number} dataStartCol - The 1-based column index where data copying starts.
 * @property {number} dataEndCol - The 1-based column index where data copying ends.
 */

/**
 * Configuration for the writable target sheet.
 * @typedef {Object} TargetConfig
 * @property {string} sheetName - The name of the target sheet.
 * @property {number} keyCol - The 1-based column index containing the unique ID to match against source.
 * @property {number} writeStartCol - The 1-based column index where the joined data will be pasted.
 */

/**
 * Joins data from a Source to a Target based on a common Key.
 * * @param {Spreadsheet} spreadsheet - The active spreadsheet object.
 * @param {SourceConfig} sourceConfig - definition of where data comes from.
 * @param {TargetConfig} targetConfig - definition of where data goes to.
 * @returns {{processed: number, updated: number}} Stats object.
 */
function joinSheetData(spreadsheet, sourceConfig, targetConfig) {
  const sourceSheet = spreadsheet.getSheetByName(sourceConfig.sheetName);
  const targetSheet = spreadsheet.getSheetByName(targetConfig.sheetName);

  if (!sourceSheet || !targetSheet) {
    throw new Error(`Missing sheets. Checked for: '${sourceConfig.sheetName}' and '${targetConfig.sheetName}'`);
  }

  // --- 1. Build Lookup Map from Source ---
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceMap = new Map();
  const numColsToCopy = sourceConfig.dataEndCol - sourceConfig.dataStartCol + 1;

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const key = row[sourceConfig.keyCol - 1]; 

    if (key && key !== "") {
      const values = row.slice(sourceConfig.dataStartCol - 1, sourceConfig.dataEndCol);
      sourceMap.set(key.toString(), values);
    }
  }

  // --- 2. Identify Rows in Target ---
  const allCheckValues = targetSheet.getRange(targetConfig.startRow, 1, targetSheet.getLastRow(), 1).getValues();
  
  let rowsToProcess = 0;
  for (let i = 0; i < allCheckValues.length; i++) {
    if (allCheckValues[i][0] === "") break;
    rowsToProcess++;
  }

  if (rowsToProcess === 0) return { processed: 0, updated: 0 };

  // --- 3. Match and Prepare Output ---
  const targetKeys = targetSheet.getRange(targetConfig.startRow, targetConfig.keyCol, rowsToProcess, 1).getValues();
  const outputData = [];
  let updateCount = 0;

  for (let i = 0; i < rowsToProcess; i++) {
    const targetKey = targetKeys[i][0].toString();

    if (sourceMap.has(targetKey)) {
      outputData.push(sourceMap.get(targetKey));
      updateCount++;
    } else {
      outputData.push(new Array(numColsToCopy).fill(""));
    }
  }

  // --- 4. Write to Target ---
  targetSheet.getRange(2, targetConfig.writeStartCol, rowsToProcess, numColsToCopy).setValues(outputData);

  return { processed: rowsToProcess, updated: updateCount };
}