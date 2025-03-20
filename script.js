document.addEventListener('DOMContentLoaded', function() {
    const loadDataBtn = document.getElementById('loadDataBtn');
    const resultsTable = document.getElementById('resultsTable');
    const resultsBody = document.getElementById('resultsBody');
    const loading = document.getElementById('loading');
    const errorDiv = document.getElementById('error');

    loadDataBtn.addEventListener('click', loadDataFromServer);

    // Function to load data from server
    function loadDataFromServer() {
        // Show loading indicator
        loading.style.display = 'block';
        resultsTable.style.display = 'none';
        errorDiv.style.display = 'none';
        
        // Fetch the Excel file from the server
        fetch('overallData.xlsx')
            .then(response => {
                if (!response.ok) {
                    throw new Error('Failed to load the Excel file from server');
                }
                return response.arrayBuffer();
            })
            .then(arrayBuffer => {
                try {
                    const data = new Uint8Array(arrayBuffer);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Assume the first sheet contains our data
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // Convert to JSON
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    if (jsonData.length === 0) {
                        showError('No data found in the Excel file');
                        return;
                    }
                    
                    // Process the data to keep only the most recent attempts
                    const processedData = processStudentData(jsonData);
                    
                    // Display the processed data
                    displayResults(processedData);
                    
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

    function processStudentData(data) {
        // Create a map to store the most recent attempt for each student-question pair
        const studentMap = new Map();
        
        // Process each record
        data.forEach(record => {
            // Expected column names from the requirements
            const studentName = record['Std Name'];
            const questionCode = record['Question Code'];
            const score = record['Score'];
            const submissionTime = record['SubmissionTime'];
            
            if (!studentName || !questionCode || score === undefined || !submissionTime) {
                console.warn('Skipping record with missing data:', record);
                return;
            }
            
            // Create a unique key for each student-question pair
            const key = `${studentName}-${questionCode}`;
            
            // Parse the submission time
            const submissionDate = new Date(submissionTime);
            
            // If this student-question pair already exists, compare submission times
            if (studentMap.has(key)) {
                const existingRecord = studentMap.get(key);
                const existingDate = new Date(existingRecord.submissionTime);
                
                // Keep only the most recent record
                if (submissionDate > existingDate) {
                    studentMap.set(key, {
                        studentName,
                        questionCode,
                        score,
                        submissionTime
                    });
                }
            } else {
                // First record for this student-question pair
                studentMap.set(key, {
                    studentName,
                    questionCode,
                    score,
                    submissionTime
                });
            }
        });
        
        // Convert the map values to array
        return Array.from(studentMap.values());
    }

    function displayResults(data) {
        // Clear previous results
        resultsBody.innerHTML = '';
        
        // Create table rows for each record
        data.forEach(record => {
            const row = document.createElement('tr');
            
            // Create cells for each column
            const nameCell = document.createElement('td');
            nameCell.textContent = record.studentName;
            
            const questionCell = document.createElement('td');
            questionCell.textContent = record.questionCode;
            
            const scoreCell = document.createElement('td');
            scoreCell.textContent = record.score;
            
            const timeCell = document.createElement('td');
            // Format the date for better readability
            const date = new Date(record.submissionTime);
            timeCell.textContent = date.toLocaleString();
            
            // Add cells to the row
            row.appendChild(nameCell);
            row.appendChild(questionCell);
            row.appendChild(scoreCell);
            row.appendChild(timeCell);
            
            // Add the row to the table
            resultsBody.appendChild(row);
        });
        
        // Hide loading indicator and show results
        loading.style.display = 'none';
        resultsTable.style.display = 'table';
    }

    function showError(message) {
        loading.style.display = 'none';
        errorDiv.style.display = 'block';
        errorDiv.textContent = message;
    }
});