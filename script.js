document.addEventListener('DOMContentLoaded', function() {
    const loadDataBtn = document.getElementById('loadDataBtn');
    const tablesContainer = document.getElementById('tablesContainer');
    const loading = document.getElementById('loading');
    const errorDiv = document.getElementById('error');

    loadDataBtn.addEventListener('click', loadExcelFile);

    // Function to load and process the Excel file
    function loadExcelFile() {
        // Show loading indicator
        loading.style.display = 'block';
        tablesContainer.innerHTML = ''; // Clear previous tables
        errorDiv.style.display = 'none';
        
        // Fetch the Excel file from the server
        fetch('overallData.xlsx')
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
                        showError('No sheets found in the Excel file');
                        return;
                    }
                    
                    // Process student data - extract the first sheet by default
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // Convert to JSON with headers
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    if (jsonData.length === 0) {
                        showError('No data found in the Excel file');
                        return;
                    }
                    
                    // Process and filter the data to keep only most recent attempts
                    const processedData = processStudentScores(jsonData);
                    
                    // Create table with the processed data
                    createScoreTable(processedData);
                    
                    // Hide loading indicator
                    loading.style.display = 'none';
                    
                } catch (error) {
                    console.error('Error processing file:', error);
                    showError('Error processing the Excel file. Please check the format.');
                }
            })
            .catch(error => {
                console.error('Error fetching file:', error);
                showError(`Error loading the file: ${error.message}`);
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

    // Function to create a table for the processed student scores
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
});