document.addEventListener('DOMContentLoaded', function() {
    const loadDataBtn = document.getElementById('loadDataBtn');
    const tablesContainer = document.getElementById('tablesContainer');
    const loading = document.getElementById('loading');
    const errorDiv = document.getElementById('error');
    
    // Add variables to store processed data
    let processedStudentData = null;
    let uniqueStudents = [];
    let uniqueClasses = []; // New variable to store unique classes
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
                    
                    // Create table selection UI after all tables are processed
                    createTableSelectionUI(Object.keys(layoutData));
                    
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
            // Normalize the student name by trimming excess whitespace and normalizing case
            const studentName = normalizeStudentName(row['Std Name']);
            const questionCode = row['Question Code'];
            const key = `${studentName}-${questionCode}`;
            const submissionTime = new Date(row['SubmissionTime']);
            
            // Extract class from student name instead of relying on a separate Class column
            const studentClass = extractClassFromName(studentName);
            
            // If this is the first attempt for this student-question pair, or it's more recent than previous attempts
            if (!mostRecentAttempts.has(key) || 
                submissionTime > mostRecentAttempts.get(key).submissionTime) {
                
                mostRecentAttempts.set(key, {
                    studentName: studentName,
                    questionCode: questionCode,
                    score: row['Score'],
                    submissionTime: submissionTime,
                    class: studentClass // Store extracted class information
                });
            }
        });
        
        // Extract unique classes
        uniqueClasses = [...new Set(Array.from(mostRecentAttempts.values()).map(item => item.class))].sort();
        
        // Convert map back to array
        return Array.from(mostRecentAttempts.values());
    }

    // Add this new function to normalize student names
    function normalizeStudentName(name) {
        if (!name) return '';
        
        // Remove extra spaces, normalize to consistent case
        return name.trim()
            .replace(/\s+/g, ' ') // Replace multiple spaces with a single space
            .replace(/\u00A0/g, ' ') // Replace non-breaking spaces with regular spaces
            .replace(/\u200B/g, ''); // Remove zero-width spaces
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
        
        // Create class filter dropdown
        const classFilterContainer = document.createElement('div');
        classFilterContainer.className = 'class-filter-container';
        
        const classLabel = document.createElement('label');
        classLabel.textContent = 'Filter by Class: ';
        classLabel.htmlFor = 'classFilter';
        classFilterContainer.appendChild(classLabel);
        
        const classSelect = document.createElement('select');
        classSelect.id = 'classFilter';
        classSelect.className = 'class-filter';
        
        // Add default "All Classes" option
        const defaultClassOption = document.createElement('option');
        defaultClassOption.value = '';
        defaultClassOption.textContent = 'All Classes';
        classSelect.appendChild(defaultClassOption);
        
        // Add options for each class (sorted)
        uniqueClasses.sort().forEach(className => {
            const option = document.createElement('option');
            option.value = className;
            option.textContent = className;
            classSelect.appendChild(option);
        });
        
        classFilterContainer.appendChild(classSelect);
        selectorSection.appendChild(classFilterContainer);
        
        // Add filter container (existing)
        const filterContainer = document.createElement('div');
        filterContainer.className = 'filter-container';

        const filterInput = document.createElement('input');
        filterInput.type = 'text';
        filterInput.id = 'studentFilter';
        filterInput.className = 'student-filter';
        filterInput.placeholder = 'Type to filter students...';

        filterContainer.appendChild(filterInput);
        selectorSection.appendChild(filterContainer);

        // Add clear button (existing)
        const clearButton = document.createElement('button');
        clearButton.type = 'button';
        clearButton.className = 'clear-filter';
        clearButton.innerHTML = '&times;';
        clearButton.title = 'Clear filter';
        clearButton.style.display = 'none';

        clearButton.addEventListener('click', function() {
            filterInput.value = '';
            filterStudents('', '', select);
            this.style.display = 'none';
        });

        filterInput.addEventListener('input', function() {
            clearButton.style.display = this.value ? 'block' : 'none';
            filterStudents(this.value.toLowerCase(), classSelect.value, select);
        });

        filterContainer.appendChild(clearButton);

        // Create select element for students
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
            option.dataset.class = getStudentClass(studentName);
            select.appendChild(option);
        });
        
        // Add event listener to handle student selection
        select.addEventListener('change', function() {
            const selectedStudent = this.value;
            if (selectedStudent) {
                updateTablesWithStudentScores(selectedStudent);
                createSummaryTable(selectedStudent); // Add this line to create summary table
            } else {
                removeStudentScoreColumns();
                createSummaryTable(null); // Add this line to remove summary table
            }
        });
        
        // Add event listener for class filter
        classSelect.addEventListener('change', function() {
            filterStudents(filterInput.value.toLowerCase(), this.value, select);
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
        
        // Create summary table with statistics for the selected student
        createSummaryTable(studentName);
        
        // Get all tables
        const tables = document.querySelectorAll('.sheet-table');
        
        tables.forEach(table => {
            const sheetName = table.closest('section').querySelector('h2').textContent;
            const sheetData = layoutData[sheetName];
            
            if (!sheetData || sheetData.length < 2) return;
            
            // Add header for score column
            const thead = table.querySelector('thead tr');
            const scoreHeader = document.createElement('th');
            
            // Extract class and number from student name (e.g., "5E 02" from "5E 02 CHEN YANTONG MARY")
            const shortName = extractClassAndNumber(studentName);
            scoreHeader.textContent = `${shortName}'s Score`;
            
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
    function filterStudents(filterText, classFilter, selectElement) {
        const options = selectElement.querySelectorAll('option');
        let visibleCount = 0;
        
        // Skip the first option (default "Select a student" option)
        for (let i = 1; i < options.length; i++) {
            const option = options[i];
            const studentName = option.textContent.toLowerCase();
            const studentClass = option.dataset.class;
            
            if ((filterText === '' || studentName.includes(filterText)) &&
                (classFilter === '' || studentClass === classFilter)) {
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
    
    // Helper function to get a student's class
    function getStudentClass(studentName) {
        // First check if we have this student in our processed data
        const studentRecord = processedStudentData.find(item => item.studentName === studentName);
        if (studentRecord && studentRecord.class) {
            return studentRecord.class;
        }
        
        // If not found in the data or class is not available, extract it from the name
        return extractClassFromName(studentName);
    }

    // Extract class from student name
    function extractClassFromName(studentName) {
        // Check if the name follows the expected format
        const nameParts = studentName.trim().split(' ');
        
        // If the first part is a class identifier (like "5E" or "5A")
        if (nameParts.length >= 1 && /^[0-9][A-F]$/.test(nameParts[0])) {
            return nameParts[0]; // Return the class part
        }
        
        return 'Unknown'; // Default return if format doesn't match
    }

    // Add this new function to extract class and student number
    function extractClassAndNumber(studentName) {
        // Match the pattern like "5E 02" at the start of the name
        const match = studentName.match(/^(\d[A-F]\s\d\d)/);
        if (match) {
            return match[1]; // Return the matched pattern
        }
        return studentName; // Fallback to the full name if pattern doesn't match
    }

    // Function to create table selection UI
    function createTableSelectionUI(sheetNames) {
        if (!sheetNames || sheetNames.length === 0) return;
        
        // Create a section for table selection
        const selectionSection = document.createElement('section');
        selectionSection.className = 'table-selection';
        selectionSection.id = 'tableSelectionSection';
        
        // Create heading
        const heading = document.createElement('h2');
        heading.textContent = 'Table Selection';
        selectionSection.appendChild(heading);
        
        // Create description
        const description = document.createElement('p');
        description.textContent = 'Select which tables to display:';
        selectionSection.appendChild(description);
        
        // Create container for checkboxes
        const checkboxContainer = document.createElement('div');
        checkboxContainer.className = 'checkbox-container';
        
        // Add "Select All" and "Deselect All" buttons
        const buttonContainer = document.createElement('div');
        buttonContainer.className = 'selection-buttons';
        
        const selectAllBtn = document.createElement('button');
        selectAllBtn.textContent = 'Select All';
        selectAllBtn.className = 'selection-btn select-all';
        selectAllBtn.addEventListener('click', function() {
            const checkboxes = checkboxContainer.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(checkbox => {
                checkbox.checked = true;
                toggleTableVisibility(checkbox.value, true);
            });
        });
        
        const deselectAllBtn = document.createElement('button');
        deselectAllBtn.textContent = 'Deselect All';
        deselectAllBtn.className = 'selection-btn deselect-all';
        deselectAllBtn.addEventListener('click', function() {
            const checkboxes = checkboxContainer.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(checkbox => {
                checkbox.checked = false;
                toggleTableVisibility(checkbox.value, false);
            });
        });
        
        buttonContainer.appendChild(selectAllBtn);
        buttonContainer.appendChild(deselectAllBtn);
        selectionSection.appendChild(buttonContainer);
        
        // Add checkboxes for each table
        sheetNames.forEach(sheetName => {
            const checkboxDiv = document.createElement('div');
            checkboxDiv.className = 'checkbox-item';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `table-${sheetName.replace(/\s+/g, '-')}`;
            checkbox.value = sheetName;
            checkbox.checked = true; // All tables visible by default
            
            checkbox.addEventListener('change', function() {
                toggleTableVisibility(this.value, this.checked);
            });
            
            const label = document.createElement('label');
            label.htmlFor = checkbox.id;
            label.textContent = sheetName;
            
            checkboxDiv.appendChild(checkbox);
            checkboxDiv.appendChild(label);
            checkboxContainer.appendChild(checkboxDiv);
        });
        
        selectionSection.appendChild(checkboxContainer);
        
        // Insert the selection section after the student selector (if present) or at the beginning
        const studentSelector = document.querySelector('.student-selector');
        if (studentSelector) {
            tablesContainer.insertBefore(selectionSection, studentSelector.nextSibling);
        } else {
            tablesContainer.insertBefore(selectionSection, tablesContainer.firstChild);
        }
    }
    
    // Function to toggle table visibility
    function toggleTableVisibility(sheetName, isVisible) {
        // Find all the h2 elements in layout-results sections
        const headings = document.querySelectorAll('.layout-results h2');
        
        // Loop through them to find the one with matching text
        for (let i = 0; i < headings.length; i++) {
            if (headings[i].textContent === sheetName) {
                const tableSection = headings[i].closest('section');
                if (tableSection) {
                    tableSection.style.display = isVisible ? '' : 'none';
                }
                break;
            }
        }
    }
    
    // Delete or comment out this function
    // Helper function to find elements by text content (for toggleTableVisibility)
    // Element.prototype.contains = function(text) {
    //     return this.textContent === text;
    // };

    // Function to create a summary table with statistics for each dataset
    function createSummaryTable(studentName) {
        // Remove any existing summary table
        const existingSummary = document.getElementById('summaryTableSection');
        if (existingSummary) {
            existingSummary.remove();
        }

        // If no student is selected, don't create a summary table
        if (!studentName) return;
        
        // Create a section for the summary table
        const summarySection = document.createElement('section');
        summarySection.className = 'results';
        summarySection.id = 'summaryTableSection';
        
        // Create heading
        const heading = document.createElement('h2');
        heading.textContent = 'Performance Summary';
        summarySection.appendChild(heading);
        
        // Create table
        const table = document.createElement('table');
        table.className = 'sheet-table summary-table';
        
        // Create table header
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        
        // Add an empty cell for the first column (metrics labels)
        const emptyHeader = document.createElement('th');
        emptyHeader.textContent = 'Metrics';
        headerRow.appendChild(emptyHeader);
        
        // Get all visible tables (only include tables that are currently visible)
        const visibleTableSections = Array.from(document.querySelectorAll('.layout-results'))
            .filter(section => section.style.display !== 'none');
        
        // Add headers for each visible table
        visibleTableSections.forEach(section => {
            const tableHeader = document.createElement('th');
            const tableName = section.querySelector('h2').textContent;
            tableHeader.textContent = tableName;
            headerRow.appendChild(tableHeader);
        });
        
        thead.appendChild(headerRow);
        table.appendChild(thead);
        
        // Create table body with rows for each metric
        const tbody = document.createElement('tbody');
        
        // Row 1: Average score
        const avgRow = document.createElement('tr');
        const avgLabel = document.createElement('td');
        avgLabel.textContent = 'Average Score';
        avgRow.appendChild(avgLabel);
        
        // Row 2: Percentage answered
        const percentRow = document.createElement('tr');
        const percentLabel = document.createElement('td');
        percentLabel.textContent = 'Completion (%)';
        percentRow.appendChild(percentLabel);
        
        // Row 3: Total questions
        const totalRow = document.createElement('tr');
        const totalLabel = document.createElement('td');
        totalLabel.textContent = 'Total Questions';
        totalRow.appendChild(totalLabel);
        
        // Calculate and add statistics for each table
        visibleTableSections.forEach(section => {
            const sheetName = section.querySelector('h2').textContent;
            const table = section.querySelector('.sheet-table');
            
            // Get all question codes from this table
            const questionCodes = [];
            const rows = table.querySelectorAll('tbody tr');
            rows.forEach(row => {
                const codeCell = row.querySelector('td:nth-child(2)');
                if (codeCell && codeCell.textContent.trim()) {
                    questionCodes.push(codeCell.textContent.trim());
                }
            });
            
            // Count total questions
            const totalQuestions = questionCodes.length;
            
            // Find student's scores for questions in this table
            let answeredQuestions = 0;
            let totalScore = 0;
            
            questionCodes.forEach(code => {
                const studentRecord = processedStudentData.find(item => 
                    item.studentName === studentName && item.questionCode === code);
                
                if (studentRecord) {
                    answeredQuestions++;
                    totalScore += parseFloat(studentRecord.score) || 0;
                }
            });
            
            // Calculate average score (prevent division by zero)
            const avgScore = answeredQuestions > 0 ? (totalScore / answeredQuestions).toFixed(2) : 'N/A';
            
            // Calculate percentage of questions answered
            const percentAnswered = totalQuestions > 0 ? 
                ((answeredQuestions / totalQuestions) * 100).toFixed(1) : '0.0';
            
            // Add cells to the rows
            const avgCell = document.createElement('td');
            avgCell.textContent = avgScore;
            if (avgScore !== 'N/A' && parseFloat(avgScore) > 0) {
                avgCell.classList.add('positive-score');
            }
            avgRow.appendChild(avgCell);
            
            const percentCell = document.createElement('td');
            percentCell.textContent = `${percentAnswered}%`;
            percentRow.appendChild(percentCell);
            
            const totalCell = document.createElement('td');
            totalCell.textContent = totalQuestions;
            totalRow.appendChild(totalCell);
        });
        
        // Add rows to table body
        tbody.appendChild(avgRow);
        tbody.appendChild(percentRow);
        tbody.appendChild(totalRow);
        table.appendChild(tbody);
        
        summarySection.appendChild(table);
        
        // Insert the summary after the table selection UI, but before the tables
        const tableSelectionSection = document.getElementById('tableSelectionSection');
        if (tableSelectionSection) {
            tablesContainer.insertBefore(summarySection, tableSelectionSection.nextSibling);
        } else {
            const studentSelector = document.querySelector('.student-selector');
            if (studentSelector) {
                tablesContainer.insertBefore(summarySection, studentSelector.nextSibling);
            } else {
                tablesContainer.appendChild(summarySection);
            }
        }
    }
});