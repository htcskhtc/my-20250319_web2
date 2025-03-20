# Student Score Tracker

A web application that loads and displays student score data from Excel files, showing only the most recent attempts for each student-question pair.

## Features

- Load student score data from Excel files (.xlsx)
- Automatically filter and display only the most recent attempts
- Clean, responsive user interface
- Client-side Excel parsing (no server-side processing required)

## Technologies Used

- HTML5
- CSS3
- JavaScript (ES6+)
- [SheetJS](https://sheetjs.com/) for Excel file parsing

## How to Use

1. Host the application on any web server
2. Ensure the `overallData.xlsx` file is in the root directory
3. Open the application in a web browser
4. Click "Load Data" to process and display the student scores

## Excel Data Format

The application expects the following columns in the Excel file:
- `Std Name` - Student name
- `Question Code` - Identifier for the question
- `Score` - Numeric score value
- `SubmissionTime` - Date/time of submission (used to determine most recent attempts)

## Project Structure

- `index.html` - Main HTML structure
- `styles.css` - CSS styling
- `script.js` - JavaScript functionality
- `overallData.xlsx` - Source data file

## License

[MIT License](LICENSE)