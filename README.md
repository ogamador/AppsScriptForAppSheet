# AppsScriptForAppSheet

AppsScriptForAppSheet is a Google Apps Script repository that provides functionality to insert images from various sources into Google Sheets cells, serving as a fix for improper image handling when generating Excel templates from AppSheet. This solution makes it easier for AppSheet users to manage images in their sheets. Additionally, it can convert Excel files to Google Sheets and supports error handling with retries, ensuring smooth operation.

## Functions

### insertImagesAndConvertToGoogleSheet

This function takes an input spreadsheet ID, sheet name, column with image URLs or IDs, and an output column. If the input file is an Excel file, it will be converted to Google Sheets. It then inserts the images into the specified output column. If any errors occur during the process, they will be logged.

### findURL

This function takes an input value and generates the appropriate URL for the image. It supports images hosted on Google Drive, Google Photos, and other sources.

### insertImage

This function takes a URL and a cell as input and inserts the image into the cell.

### convertExcelToGoogleSheet

This function takes an Excel file ID and converts the file to Google Sheets. The original Excel file is then moved to the trash.

### convertToGoogleSheet

This helper function converts an Excel file to Google Sheets and returns the new Google Sheets file.

## Usage

1. Copy the code from the Apps Script file.
2. Go to your Google Sheets file or create a new one.
3. Click on "Extensions" > "Apps Script".
4. Paste the copied code into the script editor.
5. Replace the default spreadsheet ID in the `insertImagesAndConvertToGoogleSheet` function with your own.
6. Save and run the script.

Note: Make sure to enable the Google Drive API in the script editor under "Resources" > "Advanced Google services" for the script to work properly.

