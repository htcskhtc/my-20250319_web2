document.addEventListener('DOMContentLoaded', function() {
    const loadDataBtn = document.getElementById('loadDataBtn');
    const tablesContainer = document.getElementById('tablesContainer');
    const loading = document.getElementById('loading');
    const errorDiv = document.getElementById('error');
    
    // Add variables to store processed data
    let processedStudentData = null;
    let uniqueStudents = [];
    let layoutData = {};
    
    loadDataBtn.addEventListener('click', loadExcelFiles);

    // Function to load and process both Excel files
    function loadExcelFiles() {
        // Show loading indicator
        loading.style.display = 'block';
        tablesContainer.innerHTML = ''; // Clear previous tables
        errorDiv.style.display = 'none';
        
        // Reset data storage
        processedStudentData = null;
        uniqueStudents = [];
        layoutData = {};
        
        // Create promises for both file loads
        const overallDataPromise = loadAndProcessOverallData();
        const outputLayoutPromise = loadAndProcessOutputLayout();
        
        // Process both files
        Promise.all([overallDataPromise, outputLayoutPromise])
            .then(() => {
                // Hide loading indicator when both are done
                loading.style.display = 'none';
                
                // Create student selector if we have student data
                if (processedStudentData && processedStudentData.length > 0) {
                    createStudentSelector();
                    showSuccess("Data processed successfully. Select a student to view their scores.");
                } else {
                    showError("No student data was found. Please check the data file.");
                }
            })
            .catch(error => {
                console.error('Error processing files:', error);
                showError(`Error: ${error.message}`);
            });
    }
    
    // Function to load and process the overallData.xlsx file
    function loadAndProcessOverallData() {
        return fetch('overallData.xlsx')
            .then(response => {
                if (!response.ok) {
                    throw new Error('Failed to load overallData.xlsx from server');
                }
                return response.arrayBuffer();
            })
            .then(arrayBuffer => {
                try {
                    const data = new Uint8Array(arrayBuffer);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Check if we have at least one sheet
                    if (workbook.SheetNames.length === 0) {
                        showError('No sheets found in the overallData.xlsx file');
                        return;
                    }
                    
                    // Process student data - extract the first sheet by default
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // Convert to JSON with headers
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    if (jsonData.length === 0) {
                        showError('No data found in the overallData.xlsx file');
                        return;
                    }
                    
                    // Process and filter the data to keep only most recent attempts
                    processedStudentData = processStudentScores(jsonData);
                    
                    // Extract unique student names
                    uniqueStudents = [...new Set(processedStudentData.map(item => item.studentName))].sort();
                    
                    console.log(`Processed ${processedStudentData.length} student records with ${uniqueStudents.length} unique students`);
                } catch (error) {
                    console.error('Error processing overallData.xlsx:', error);
                    showError('Error processing the overallData.xlsx file. Please check the format.');
                    throw error;
                }
            });
    }
    
    // Function to load and process the outputLayout.xlsx file
    function loadAndProcessOutputLayout() {
        return fetch('outputLayout.xlsx')
            .then(response => {
                if (!response.ok) {
                    throw new Error('Failed to load outputLayout.xlsx from server');
                }
                return response.arrayBuffer();
            })
            .then(arrayBuffer => {
                try {
                    const data = new Uint8Array(arrayBuffer);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Check if we have sheets
                    if (workbook.SheetNames.length === 0) {
                        console.warn('No sheets found in the outputLayout.xlsx file');
                        return;
                    }
                    
                    // Process each sheet in the workbook and store the data
                    workbook.SheetNames.forEach(sheetName => {
                        const worksheet = workbook.Sheets[sheetName];
                        
                        // Convert to JSON with headers (using header: 1 to get array of arrays)
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        if (jsonData.length === 0) {
                            console.warn(`No data found in sheet: ${sheetName}`);
                            return;
                        }
                        
                        // Store layout data for later use
                        layoutData[sheetName] = jsonData;
                        
                        // Create a table for this sheet (showing only first two columns)
                        createLayoutTable(sheetName, jsonData);
                    });
                    
                } catch (error) {
                    console.error('Error processing outputLayout.xlsx:', error);
                    showError('Error processing the outputLayout.xlsx file. Please check the format.');
                    throw error;
                }
            });
    }

    // Function to process student scores and keep only most recent attempts
    function processStudentScores(data) {
        // Check if data has the expected columns
        if (!data.length || !('Std Name' in data[0]) || !('Question Code' in data[0]) || 
            !('Score' in data[0]) || !('SubmissionTime' in data[0])) {
            throw new Error('Data format is not valid. Expected columns: Std Name, Question Code, Score, SubmissionTime');
        }
        
        // Create a map to store the most recent attempt for each student-question pair
        const mostRecentAttempts = new Map();
        
        // Process each row of data
        data.forEach(row => {
            const studentName = row['Std Name'];
            const questionCode = row['Question Code'];
            const key = `${studentName}-${questionCode}`;
            const submissionTime = new Date(row['SubmissionTime']);
            
            // If this is the first attempt for this student-question pair, or it's more recent than previous attempts
            if (!mostRecentAttempts.has(key) || 
                submissionTime > mostRecentAttempts.get(key).submissionTime) {
                
                mostRecentAttempts.set(key, {
                    studentName: studentName,
                    questionCode: questionCode,
                    score: row['Score'],
                    submissionTime: submissionTime
                });
            }
        });
        
        // Convert map back to array
        return Array.from(mostRecentAttempts.values());
    }
    
    // Function to create a dropdown selector for students
    function createStudentSelector() {
        // Create a section for the student selector
        const selectorSection = document.createElement('section');
        selectorSection.className = 'student-selector';
        
        // Create heading
        const heading = document.createElement('h2');
        heading.textContent = 'Select a Student';
        selectorSection.appendChild(heading);
        
        // Add this inside the createStudentSelector function, after creating the heading
        const filterContainer = document.createElement('div');
        filterContainer.className = 'filter-container';

        const filterInput = document.createElement('input');
        filterInput.type = 'text';
        filterInput.id = 'studentFilter';
        filterInput.className = 'student-filter';
        filterInput.placeholder = 'Type to filter students...';

        filterContainer.appendChild(filterInput);
        selectorSection.appendChild(filterContainer);

        // Add this after creating the filter input
        const clearButton = document.createElement('button');
        clearButton.type = 'button';
        clearButton.className = 'clear-filter';
        clearButton.innerHTML = '&times;';
        clearButton.title = 'Clear filter';
        clearButton.style.display = 'none';

        clearButton.addEventListener('click', function() {
            filterInput.value = '';
            filterStudents('', select);
            this.style.display = 'none';
        });

        filterInput.addEventListener('input', function() {
            clearButton.style.display = this.value ? 'block' : 'none';
            filterStudents(this.value.toLowerCase(), select);
        });

        filterContainer.appendChild(clearButton);

        // Create select element
        const select = document.createElement('select');
        select.id = 'studentSelect';
        select.className = 'student-select';
        
        // Add default option
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = '-- Select a student --';
        select.appendChild(defaultOption);
        
        // Add an option for each student
        uniqueStudents.forEach(studentName => {
            const option = document.createElement('option');
            option.value = studentName;
            option.textContent = studentName;
            select.appendChild(option);
        });
        
        // Add event listener to handle student selection
        select.addEventListener('change', function() {
            const selectedStudent = this.value;
            if (selectedStudent) {
                updateTablesWithStudentScores(selectedStudent);
            } else {
                removeStudentScoreColumns();
            }
        });
        
        // Add this before appending the select element
        filterInput.addEventListener('input', function() {
            filterStudents(this.value.toLowerCase(), select);
        });

        selectorSection.appendChild(select);
        
        // Insert the selector at the beginning of tablesContainer
        if (tablesContainer.firstChild) {
            tablesContainer.insertBefore(selectorSection, tablesContainer.firstChild);
        } else {
            tablesContainer.appendChild(selectorSection);
        }
    }
    
    // Function to update tables with selected student scores
    function updateTablesWithStudentScores(studentName) {
        // First remove any existing score columns
        removeStudentScoreColumns();
        
        // Get all tables
        const tables = document.querySelectorAll('.sheet-table');
        
        tables.forEach(table => {
            const sheetName = table.closest('section').querySelector('h2').textContent;
            const sheetData = layoutData[sheetName];
            
            if (!sheetData || sheetData.length < 2) return;
            
            // Add header for score column
            const thead = table.querySelector('thead tr');
            const scoreHeader = document.createElement('th');
            scoreHeader.textContent = `${studentName}'s Score`;
            scoreHeader.className = 'student-score-column';
            thead.appendChild(scoreHeader);
            
            // Get question codes from the table (assuming they are in the second column)
            const tbody = table.querySelector('tbody');
            const rows = tbody.querySelectorAll('tr');
            
            rows.forEach((row, index) => {
                // Get question code from the row (second column)
                const questionCodeCell = row.querySelector('td:nth-child(2)');
                if (!questionCodeCell) return;
                
                const questionCode = questionCodeCell.textContent.trim();
                if (!questionCode) return;
                
                // Find the student's score for this question
                const studentRecord = processedStudentData.find(item => 
                    item.studentName === studentName && item.questionCode === questionCode);
                
                // Add a cell for the score
                const scoreCell = document.createElement('td');
                scoreCell.className = 'student-score-column';
                
                if (studentRecord) {
                    scoreCell.textContent = studentRecord.score;
                    // Add a class based on the score value for styling
                    if (studentRecord.score > 0) {
                        scoreCell.classList.add('positive-score');
                    } else {
                        scoreCell.classList.add('zero-score');
                    }
                } else {
                    scoreCell.textContent = 'N/A';
                    scoreCell.classList.add('no-score');
                }
                
                row.appendChild(scoreCell);
            });
        });
    }
    
    // Function to remove student score columns from all tables
    function removeStudentScoreColumns() {
        const scoreColumns = document.querySelectorAll('.student-score-column');
        scoreColumns.forEach(column => column.remove());
    }

    // Function to create a table for the layout data
    function createLayoutTable(sheetName, data) {
        // Create a section for this table
        const tableSection = document.createElement('section');
        tableSection.className = 'results layout-results';
        
        // Create heading with sheet name
        const heading = document.createElement('h2');
        heading.textContent = sheetName;
        tableSection.appendChild(heading);
        
        // Create table
        const table = document.createElement('table');
        table.className = 'sheet-table';
        table.dataset.sheetName = sheetName;
        
        // Create table header
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        
        // Use the first row as header, but only the first two columns
        if (data[0] && data[0].length > 0) {
            // Create headers for first two columns
            for (let i = 0; i < Math.min(2, data[0].length); i++) {
                const th = document.createElement('th');
                th.textContent = data[0][i] || `Column ${i+1}`;
                headerRow.appendChild(th);
            }
        } else {
            // Default headers if first row is empty
            const th1 = document.createElement('th');
            th1.textContent = 'Column 1';
            headerRow.appendChild(th1);
            
            const th2 = document.createElement('th');
            th2.textContent = 'Column 2';
            headerRow.appendChild(th2);
        }
        
        thead.appendChild(headerRow);
        table.appendChild(thead);
        
        // Create table body
        const tbody = document.createElement('tbody');
        
        // Add data rows (skip the header row)
        for (let i = 1; i < data.length; i++) {
            const row = document.createElement('tr');
            
            // Only display the first two columns
            for (let j = 0; j < Math.min(2, data[i].length); j++) {
                const cell = document.createElement('td');
                cell.textContent = data[i][j] !== undefined ? data[i][j] : '';
                row.appendChild(cell);
            }
            
            // If the row has fewer than 2 columns, add empty cells
            for (let j = data[i].length; j < 2; j++) {
                const cell = document.createElement('td');
                cell.textContent = '';
                row.appendChild(cell);
            }
            
            tbody.appendChild(row);
        }
        
        table.appendChild(tbody);
        tableSection.appendChild(table);
        
        // Add the table section to the container
        tablesContainer.appendChild(tableSection);
    }

    // Function to create a table for the processed student scores (kept for backward compatibility)
    function createScoreTable(processedData) {
        // Create a section for the table
        const tableSection = document.createElement('section');
        tableSection.className = 'results';
        
        // Create heading
        const heading = document.createElement('h2');
        heading.textContent = 'Student Scores (Most Recent Attempts)';
        tableSection.appendChild(heading);
        
        // Create table
        const table = document.createElement('table');
        table.className = 'score-table';
        
        // Create table header
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        
        // Add headers
        const headers = ['Student Name', 'Question Code', 'Score', 'Submission Time'];
        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });
        
        thead.appendChild(headerRow);
        table.appendChild(thead);
        
        // Create table body
        const tbody = document.createElement('tbody');
        
        // Sort the data by student name and then by question code
        processedData.sort((a, b) => {
            // First sort by student name
            if (a.studentName < b.studentName) return -1;
            if (a.studentName > b.studentName) return 1;
            
            // If same student, sort by question code
            return a.questionCode.localeCompare(b.questionCode);
        });
        
        // Add data rows
        processedData.forEach(item => {
            const row = document.createElement('tr');
            
            // Create cells
            const nameCell = document.createElement('td');
            nameCell.textContent = item.studentName;
            row.appendChild(nameCell);
            
            const questionCell = document.createElement('td');
            questionCell.textContent = item.questionCode;
            row.appendChild(questionCell);
            
            const scoreCell = document.createElement('td');
            scoreCell.textContent = item.score;
            row.appendChild(scoreCell);
            
            const timeCell = document.createElement('td');
            timeCell.textContent = item.submissionTime.toLocaleString();
            row.appendChild(timeCell);
            
            tbody.appendChild(row);
        });
        
        table.appendChild(tbody);
        tableSection.appendChild(table);
        
        // Add the table section to the container
        tablesContainer.appendChild(tableSection);
    }

    function showError(message) {
        loading.style.display = 'none';
        errorDiv.style.display = 'block';
        errorDiv.textContent = message;
    }

    function showSuccess(message) {
        loading.style.display = 'none';
        errorDiv.style.display = 'block';
        errorDiv.textContent = message;
    }

    // Function to filter students based on input text
    function filterStudents(filterText, selectElement) {
        const options = selectElement.querySelectorAll('option');
        let visibleCount = 0;
        
        // Skip the first option (default "Select a student" option)
        for (let i = 1; i < options.length; i++) {
            const option = options[i];
            const studentName = option.textContent.toLowerCase();
            
            if (filterText === '' || studentName.includes(filterText)) {
                option.style.display = '';
                visibleCount++;
            } else {
                option.style.display = 'none';
            }
        }
        
        // Show message if no students match
        const existingMessage = document.getElementById('noMatchMessage');
        if (visibleCount === 0 && filterText !== '') {
            if (!existingMessage) {
                const message = document.createElement('div');
                message.id = 'noMatchMessage';
                message.className = 'no-match-message';
                message.textContent = 'No students match your filter';
                selectElement.parentNode.appendChild(message);
            }
        } else if (existingMessage) {
            existingMessage.remove();
        }
    }
});