let excelData = [];
let filteredData = [];

// Handle file upload and display data
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(event) {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
            displayTable(excelData);
        };
        reader.readAsArrayBuffer(file);
    }
});

// Function to display the table data (focused on student results)
function displayTable(data) {
    const tableHeader = document.getElementById('tableHeader');
    const tableBody = document.getElementById('tableBody');
    tableHeader.innerHTML = '';
    tableBody.innerHTML = '';

    if (data.length > 0) {
        // Create headers
        const headers = data[0];  // Assuming first row is headers
        headers.forEach(header => {
            const th = document.createElement('th');
            th.innerText = header;
            tableHeader.appendChild(th);
        });

        // Populate rows with student data
        data.slice(1).forEach(row => {
            const tr = document.createElement('tr');
            row.forEach(cell => {
                const td = document.createElement('td');
                td.innerText = cell;
                tr.appendChild(td);
            });
            tableBody.appendChild(tr);
        });
    }
}

// Handle search for student data
document.getElementById('searchInput').addEventListener('input', function() {
    const searchValue = this.value.toLowerCase();
    filteredData = excelData.filter((row, index) => {
        if (index === 0) return false;  // Skip the header row
        return row.some(cell => String(cell).toLowerCase().includes(searchValue));
    });
    displayTable(filteredData.length ? [excelData[0], ...filteredData] : excelData); // Include header in filtered data
    document.getElementById('downloadButton').disabled = filteredData.length === 0;
});

// Handle download of filtered student results
document.getElementById('downloadButton').addEventListener('click', function() {
    if (filteredData.length > 0) {
        const ws = XLSX.utils.aoa_to_sheet([excelData[0], ...filteredData]); // Include headers
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "FilteredStudentResults");
        XLSX.writeFile(wb, "filtered_student_results.xlsx");
    }
});
