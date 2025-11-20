/**
 * Generates slides based on the provided configuration object.
 * * @param {Object} config - The configuration settings containing:
 * * @param {string} config.sourceSheetName - The name of the sheet tab to read from (e.g., 'Dashboard').
 * * @param {string[]} config.dataRanges - An array of A1 notation ranges to process (e.g., ['C102:O122']).
 * * @param {string} config.targetPresentationId - The ID string of the Google Slides presentation.
 * * @param {number} config.lowSatTemplateIndex  - The zero-based index of the LowSat template slide (Page Number - 1).
 * * @param {number} config.escalationTemplateIndex - The zero-based index of the Escalation template slide (Page Number - 1).
 * * @param {string} [config.sourceSheetId] - The ID of the spreadsheet. Defaults to the active spreadsheet if omitted.
 */

function generateSlides(config) {
  const ui = SpreadsheetApp.getUi();
  try {
    // --- 1. Validate Configuration and Get Sheet ---
    if (!config.sourceSheetName || !config.dataRanges || !config.targetPresentationId || config.templateSlideIndex == null) {
      throw new Error("Configuration object is missing one or more required properties.");
    }
    const sourceSheetId = config.sourceSheetId || SpreadsheetApp.getActiveSpreadsheet().getId();
    const spreadsheet = SpreadsheetApp.openById(sourceSheetId);
    const sheet = spreadsheet.getSheetByName(config.sourceSheetName);
    if (!sheet) {
      throw new Error(`Sheet "${config.sourceSheetName}" not found in spreadsheet ID "${sourceSheetId}".`);
    }

    // --- 2. Access Presentation and Template Slide ---
    const presentation = SlidesApp.openById(config.targetPresentationId);
    const slides = presentation.getSlides();
    if (config.templateSlideIndex >= slides.length) {
       throw new Error(`Template slide index ${config.templateSlideIndex} is out of bounds. The presentation only has ${slides.length} slides.`);
    }
    const templateSlide = slides[config.templateSlideIndex];
    let totalSlidesCreated = 0;

    // --- 3. Loop Through Data and Create Slides ---
    // We process ranges from last to first to handle multi-range cases correctly.
    for (let i = config.dataRanges.length - 1; i >= 0; i--) {
      const range = config.dataRanges[i];
      console.log(`Processing data from range: ${range}`);
      const fullRangeData = sheet.getRange(range).getValues();
      const headers = fullRangeData.shift(); // First row is headers
      const dataRows = fullRangeData;

      for (let j = dataRows.length - 1; j >= 0; j--) {
        const row = dataRows[j];

        // Skip empty rows, but continue checking upwards in the range.
        if (!row[0] || row[0].toString().trim() === '') {
          continue;
        }
        const newSlide = templateSlide.duplicate();
        
        console.log(`Created new slide for row ${j + 1} in range ${range}.`);

        // Replace all standard placeholders
        headers.forEach((header, colIndex) => {
          if (header && header.trim() !== '') {
            const placeholder = `{{${header}}}`;
            const value = row[colIndex] || ''; // Use empty string for blank cells
            newSlide.replaceAllText(placeholder, value);
            console.log(`Replacing "${placeholder}" with "${value}"`);
          }
        });

        // Handle custom composite placeholders
        const caseNumberIndex = headers.indexOf('Case Number');
        const titleIndex = headers.indexOf('Title');
        if (caseNumberIndex !== -1 && titleIndex !== -1) {
          const caseNumber = row[caseNumberIndex];
          const title = row[titleIndex];
          if (caseNumber && title) {
            newSlide.replaceAllText('{{Case Title}}', `Case: ${caseNumber} ${title}`);
            console.log(`Replacing composite "{{Case Title}}" with "Case: ${caseNumber} ${title}"`);
          }
        }
        totalSlidesCreated++;
      }
    }

    // --- 4. Final User Notification ---
    if (totalSlidesCreated > 0) {
      ui.alert('Success!', `Successfully created ${totalSlidesCreated} new slides in the correct order.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Process Complete', 'No data found to generate slides.', ui.ButtonSet.OK);
    }

  } catch (error) {
    console.error(`An error occurred: ${error.message} \nStack: ${error.stack}`);
    ui.alert('Error', `An error occurred: ${error.message}`, ui.ButtonSet.OK);
  }
}

