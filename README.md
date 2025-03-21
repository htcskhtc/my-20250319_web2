# Student Score Tracker

A web application that loads and displays student score data from Excel files, showing only the most recent attempts for each student-question pair. The application allows you to select individual students and view their scores across different question sets.

## Features

- Load student score data from Excel files (.xlsx)
- Automatically filter and display only the most recent attempts for each student-question pair
- View organized question sets based on layout information
- Filter students by class and name
- Select individual students to view their specific scores
- Responsive design that works on various screen sizes
- Client-side Excel parsing (no server-side processing required)

## Technologies Used

- HTML5
- CSS3
- JavaScript (ES6+)
- [SheetJS](https://sheetjs.com/) for Excel file parsing

## How to Use

1. Prepare your Excel files:
   - `overallData.xlsx` - Contains the student score data
   - `outputLayout.xlsx` - Contains the layout information for question sets
2. Host the application on any web server (or use GitHub Pages)
3. Open the application in a web browser
4. Click "Import Excel" to process and display the data
5. Use the class filter and search box to find specific students
6. Select a student to view their scores across all question sets

## Excel Data Format

### overallData.xlsx
The application expects the following columns:
- `Std Name` - Student name with class identifier (e.g., "5A John Smith")
- `Question Code` - Identifier for the question
- `Score` - Numeric score value
- `SubmissionTime` - Date/time of submission (used to determine most recent attempts)

### outputLayout.xlsx
This file should contain multiple sheets, each representing a question set:
- First row contains column headers
- First column typically contains question descriptions
- Second column contains question codes that match those in overallData.xlsx

## Project Structure

- `index.html` - Main HTML structure and user interface
- `styles.css` - CSS styling and responsive design
- `script.js` - JavaScript functionality for data processing and display
- `overallData.xlsx` - Source data file for student scores
- `outputLayout.xlsx` - Layout information for question sets

## License

[MIT License](LICENSE)