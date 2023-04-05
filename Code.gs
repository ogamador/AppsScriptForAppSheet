/**
 * Main function that inserts images into Google Sheets and converts Excel files to Google Sheets if needed.
 * @param {string} spreadsheetID - The ID of the spreadsheet containing the images' URLs or IDs.
 * @param {string} sheetName - The name of the sheet where the images will be inserted.
 * @param {number} columnWithFoto - The column containing the image URLs or IDs.
 * @param {number} columnOutput - The column where the images will be inserted.
 */
function insertImagesAndConvertToGoogleSheet(spreadsheetID = 'Your_excel_id[_ID]', sheetName, columnWithFoto = 1, columnOutput = columnWithFoto) {
  const file = DriveApp.getFileById(spreadsheetID);
  const mimeType = file.getMimeType();

  if (mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mimeType === 'application/vnd.ms-excel') {
    spreadsheetID = convertExcelToGoogleSheet(spreadsheetID);
  } else if (mimeType !== 'application/vnd.google-apps.spreadsheet') {
    throw new Error('Unsupported file type');
  }

  const errors = [];
  let attempt = 0;

  while (attempt < 5) {
    attempt++;
    try {
      console.log("Attempt: " + attempt);
      const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
      const sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getSheets()[0];
      const values = sheet.getRange(2, columnWithFoto, sheet.getLastRow() - 1, 1).getValues().flat();
      for (let i = 0; i < values.length; i++) {
        try {
          const url = findURL(values[i]);
          if (url !== '') {
            insertImage(url, sheet.getRange(2 + i, columnOutput));
          }
        } catch (error) {
          errors.push({item: values[i], error: error.message});
        }
      }
      console.log('Function completed');
      break;
    } catch (e) {
      console.log(e);
      Utilities.sleep(3000);
    }
  }

  if (errors.length > 0) {
    console.log('Errors occurred while processing the following items:');
    errors.forEach(error => console.log(`Item: ${error.item}, Error: ${error.error}`));
  } else {
    console.log('All items were processed successfully');
  }
}

/**
 * Finds the appropriate URL for the image based on the input value.
 * @param {string} value - The input value that may contain the image URL or ID.
 * @return {string} - The URL for the image.
 */
function findURL(value) {
  if (value === '') {
    return '';
  } else if (value.slice(0, 11).toLowerCase() === 'https://lh3') {
    return value;
  } else if (value.slice(0, 5).toLowerCase() === 'https') {
    const idFromUrl = value.split("=")[2];
    const file = DriveApp.getFileById(idFromUrl);
    const blob = file.getBlob();
    const base64Image = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return `data:${mimeType};base64,${base64Image}`;
  } else if (value.split("/").length === 2) {
    const fileName = value.split("/")[1];
    const file = DriveApp.getFilesByName(fileName).next();
    const blob = file.getBlob();
    const base64Image = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return `data:${mimeType};base64,${base64Image}`;
  } else {
    return '';
  }
}

/**
 * Inserts an image into a specified cell using the image URL.
 * @param {string} url - The URL of the image to insert.
 * @param {Object} cell - The cell object where the image will be inserted.
 */
function insertImage(url, cell) {
  const image = SpreadsheetApp
    .newCellImage()
    .setSourceUrl(url)
    .setAltTextTitle('item')
    .setAltTextDescription('item')
    .build();
  cell.setValue(image);
}

/**
 * Converts an Excel file to Google Sheets and returns the new Google Sheets file's ID.
 * @param {string} excelFileId - The ID of the Excel file to convert.
 * @return {string} - The ID of the new Google Sheets file.
 */
function convertExcelToGoogleSheet(excelFileId = '1087uKPtyGTiEP-3YzssqvOtzR1q-l0v6') {
  const excelFile = DriveApp.getFileById(excelFileId);
  const googleSheet = convertToGoogleSheet(excelFile);
  excelFile.setTrashed(true);

  Logger.log(`Excel file converted to Google Sheet: Name: ${googleSheet.getName()}, ID: ${googleSheet.getId()}, URL: ${googleSheet.getUrl()}`);
  return googleSheet.getId();
}

/**
 * Helper function that converts an Excel file to Google Sheets and returns the new Google Sheets file.
 * @param {Object} excelFile - The Excel file object to convert.
 * @return {Object} - The new Google Sheets file object.
 */
function convertToGoogleSheet(excelFile) {
  const excelBlob = excelFile.getBlob();
  const parentFolder = excelFile.getParents().next();

  const resource = {
    mimeType: 'application/vnd.google-apps.spreadsheet',
    title: excelFile.getName(),
    parents: [{ id: parentFolder.getId() }]
  };

  const convertedGoogleSheet = Drive.Files.insert(resource, excelBlob);
  console.log(convertedGoogleSheet.id);

  const googleSheet = SpreadsheetApp.openById(convertedGoogleSheet.id);
  return googleSheet;
}
