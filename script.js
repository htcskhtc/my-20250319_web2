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
        fetch('outputLayout.xlsx')
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
                        showError('No sheets found in the Excel file');
                        return;
                    }
                    
                    // Process each sheet in the workbook
                    processAllSheets(workbook);
                    
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

    function processAllSheets(workbook) {
        // Process each sheet and create tables
        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            
            // Convert to JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonData.length === 0) {
                console.warn(`No data found in sheet: ${sheetName}`);
                return;
            }
            
            // Create a table for this sheet
            createSheetTable(sheetName, jsonData);
        });
        
        // Hide loading indicator
        loading.style.display = 'none';
    }

    function createSheetTable(sheetName, data) {
        // Create a section for this table
        const tableSection = document.createElement('section');
        tableSection.className = 'results';
        
        // Create heading with sheet name
        const heading = document.createElement('h2');
        heading.textContent = sheetName;
        tableSection.appendChild(heading);
        
        // Create table
        const table = document.createElement('table');
        table.className = 'sheet-table';
        
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

    function showError(message) {
        loading.style.display = 'none';
        errorDiv.style.display = 'block';
        errorDiv.textContent = message;
    }
});