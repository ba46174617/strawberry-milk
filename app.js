document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('fileInput');
    fileInput.addEventListener('change', function() {
        if (fileInput.files.length > 0) {
            const file = fileInput.files[0];
            if (file.name.endsWith('.xlsx')) {
                parseExcel(file);
            } else {
                alert('Please upload a valid Excel file.');
            }
        }
    });

    // Remove the click listener from the 'Upload Excel' button
    // Instead, let the file selection trigger the parsing directly
});

function submitToSharePoint() {
    const rows = document.querySelectorAll('#dataTable tbody tr');
    rows.forEach(row => {
        const LM = row.cells[0].querySelector('select').value;
        const baseMobilePostpaid = row.cells[1].querySelector('input').value;
        const baseMobilePrepaid = row.cells[2].querySelector('input').value;
        const baseFixed = row.cells[3].querySelector('input').value;
        const baseConsumer = row.cells[4].querySelector('input').value;
        const baseEnterprise = row.cells[5].querySelector('input').value;

        // Construct the data object to send to SharePoint
        const data = {
            Title: LM,  // 'Title' because SharePoint requires this field by default
            Base_Mobile_Postpaid: baseMobilePostpaid,
            Base_Mobile_Prepaid: baseMobilePrepaid,
            Base_Fixed: baseFixed,
            Base_Consumer: baseConsumer,
            Base_Enterprise: baseEnterprise
        };

        // AJAX call to SharePoint's REST API
        fetch('https://vodafone.sharepoint.com/sites/LMSubmissionTestSite/Lists/Test%20Submission/AllItems.aspx/_api/web/lists/getbytitle(\'Test Submission\')/items', {
            method: 'POST',
            body: JSON.stringify(data),
            headers: {
                "Accept": "application/json; odata=verbose",
                "Content-Type": "application/json; odata=verbose",
                "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value,
                "X-HTTP-Method": "MERGE",
                "If-Match": "*"
            }
        })
        .then(response => response.json())
        .then(data => console.log('Success:', data))
        .catch((error) => {
            console.error('Error:', error);
        });
    });
}

// Ensure you handle authentication and include a valid request digest or authentication headers

function addRow() {
    const table = document.getElementById('dataTable');
    const newRow = table.insertRow(-1);
    newRow.innerHTML = `
        <td>
            <select name="LM">
                <option value="">Select Local Market</option>
                <option value="RO">RO</option>
                <option value="IT">IT</option>
                <option value="ES">ES</option>
                <option value="TR">TR</option>
                <option value="DE">DE</option>
                <option value="IE">IE</option>
                <option value="PT">PT</option>
                <option value="UK">UK</option>
            </select>
        </td>
        <td><input type="number" name="Base-Mobile Postpaid" min="1" oninput="validateInteger(this)"></td>
        <td><input type="number" name="Base-Mobile Prepaid" min="1" oninput="validateInteger(this)"></td>
        <td><input type="number" name="Base-Fixed" min="1" oninput="validateInteger(this)"></td>
        <td><input type="number" name="Base-Consumer" min="1" oninput="validateInteger(this)"></td>
        <td><input type="number" name="Base-Enterprise" min="1" oninput="validateInteger(this)"></td>
        <td><button onclick="removeRow(this)">Remove</button></td>
    `;
}

function validateInteger(input) {
    if (input.value < 1 || !Number.isInteger(parseFloat(input.value))) {
        input.value = '';
        alert('Please enter an integer greater than zero.');
    }
}

function removeRow(button) {
    const row = button.parentNode.parentNode;
    row.parentNode.removeChild(row);
}

function handleFile(file) {
    console.log("Handling file:", file); // Check if the file object is received
    if (file && file.name.endsWith('.xlsx')) {
        parseExcel(file);
    } else {
        console.log("File is not valid:", file); // Log if the file is not valid
        alert('Please upload a valid Excel file.'); // This alert should only show if file isn't an Excel
    }
}

function parseExcel(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            const workbook = XLSX.read(data, {type: 'array'});
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            if (validateExcelData(json)) {
                populateTable(json);
            }
        } catch (error) {
            console.error("Error reading Excel data:", error);
            alert('Failed to read Excel file. Please check the file format and content.');
        }
    };
    reader.onerror = function(error) {
        console.error("Error reading file:", error);
        alert('Error reading file.');
    };
    reader.readAsArrayBuffer(file);
}

function validateExcelData(data) {
    const validOptions = ["RO", "IT", "ES", "TR", "DE", "IE", "PT", "UK"];
    let isValid = true;
    let errors = [];

    data.forEach((row, index) => {
        if (index === 0) return; // Skip header

        let rowErrors = [];
        if (!validOptions.includes(row[0])) {
            rowErrors.push(`Column A: '${row[0]}' is not a valid option`);
        }
        for (let i = 1; i < row.length; i++) {
            if (!(Number.isInteger(row[i]) && row[i] > 0)) {
                rowErrors.push(`Column ${String.fromCharCode(65 + i)}: '${row[i]}' is not a valid integer greater than 0`);
            }
        }

        if (rowErrors.length > 0) {
            errors.push(`Row ${index + 1}: ${rowErrors.join(", ")}`);
            isValid = false;
        }
    });

    if (!isValid) {
        alert("Errors in Excel file:\n" + errors.join("\n"));
    }
    return isValid;
}

function populateTable(data) {
    const table = document.getElementById('dataTable');
    table.getElementsByTagName('tbody')[0].innerHTML = ''; // Clear existing rows

    data.forEach((row, index) => {
        if (index === 0) return; // Skip header
        const newRow = table.insertRow(-1);
        newRow.innerHTML = `
            <td>
                <select name="LM" value="${row[0]}">
                    ${["RO", "IT", "ES", "TR", "DE", "IE", "PT", "UK"]
                    .map(option => `<option value="${option}" ${option === row[0] ? 'selected' : ''}>${option}</option>`).join('')}
                </select>
            </td>
            <td><input type="number" name="Base-Mobile Postpaid" min="1" value="${row[1]}" oninput="validateInteger(this)"></td>
            <td><input type="number" name="Base-Mobile Prepaid" min="1" value="${row[2]}" oninput="validateInteger(this)"></td>
            <td><input type="number" name="Base-Fixed" min="1" value="${row[3]}" oninput="validateInteger(this)"></td>
            <td><input type="number" name="Base-Consumer" min="1" value="${row[4]}" oninput="validateInteger(this)"></td>
            <td><input type="number" name="Base-Enterprise" min="1" value="${row[5]}" oninput="validateInteger(this)"></td>
            <td><button onclick="removeRow(this)">Remove</button></td>
        `;
    });
}