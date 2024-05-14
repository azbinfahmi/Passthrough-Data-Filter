document.getElementById('csvFileInput').addEventListener('change', handleFileSelect);

function handleFileSelect(event) {
    const fileInput = event.target;
    const file = fileInput.files[0];
    
    if (file) {
        const reader = new FileReader();

        reader.onload = function (e) {
            const csvData = parseCSV(e.target.result);
            filterAndCreateSheets(csvData);
        };

        reader.readAsText(file);
    }
}

function parseCSV(csvText) {
    const lines = csvText.split('\n');
    const header = lines[0].split(',');

    const YcoordIndex = header.indexOf('Y');
    const XcoordIndex = header.indexOf('X');
    
    return lines.slice(1).map(line => {
        const cells = line.split(',');
        
        // Extract Y (latitude) and X (longitude) coordinates from the current row
        const Ycoord = cells[YcoordIndex];
        const Xcoord = cells[XcoordIndex];
        
        // Construct Google Maps link using the coordinates from the current row
        const Gmaps = `https://www.google.com/maps?q=${Ycoord},${Xcoord}`;

        // Extract "Unit Number" and "House Number" values and remove double quotes
        const unitNumberIndex = header.indexOf('Unit Number');
        const houseNumberIndex = header.indexOf('House Number');

        let unitNumber, houseNumber
        
        if(cells[unitNumberIndex] == undefined || cells[houseNumberIndex] == undefined){
            if(cells[unitNumberIndex] == undefined){
                unitNumber = cells[unitNumberIndex]
            }
            else{
                houseNumber = cells[houseNumberIndex]
            }
            
        }

        else{
            if(cells[unitNumberIndex].includes("")){
                unitNumber = unitNumberIndex !== -1 ? cells[unitNumberIndex]?.replace(/"/g, '').trim() : '';
            }
            else{
                unitNumber = cells[unitNumberIndex]
            }

            if(cells[houseNumberIndex].includes("")){
                houseNumber = houseNumberIndex !== -1 ? cells[houseNumberIndex]?.replace(/"/g, '').trim() : ''
            }
            else{
                houseNumber = cells[houseNumberIndex]
            }
        }
        console.log('unitNumber: ',unitNumber)
        console.log('houseNumber: ',houseNumber)
        // Return the parsed data object
        return {
            Y: Ycoord, 
            X: Xcoord, 
            'GoogleMaps': Gmaps,
            v_layer: cells[header.indexOf('v_layer')],
            v_project: cells[header.indexOf('v_project')],
            ID: cells[header.indexOf('ID')],
            'Unit Number': unitNumber,
            'House Number': houseNumber,
            Street: cells[header.indexOf('Street')],
            City: cells[header.indexOf('City')],
            'Bldg Type Source': cells[header.indexOf('Bldg Type Source')],
            Note: cells[header.indexOf('Note')],
            Subname: cells[header.indexOf('Subname')],
            County: cells[header.indexOf('County')],
            'Street Address': cells[header.indexOf('Street Address')],
            'Building Type': cells[header.indexOf('Building Type')],
            'EBB Remark': '',
            'Action': '',
            'Producer': '',
            'Date': ''
        };
    });
}

function filterAndCreateSheets(csvData) {
    // Create a workbook
    const wb = XLSX.utils.book_new();

    // Ask the user for a custom file name
    const customFileName = prompt("Enter the file name:") + " .xlsx";
    const fileName = customFileName ? customFileName : 'Addressing.xlsx';

    // Convert data to sheet
    const ws = XLSX.utils.json_to_sheet(csvData.map(row => {
        // Create a new object to hold the data for this row
        const newRow = {};
        
        // Copy data from the original row
        Object.keys(row).forEach(key => {
            newRow[key] = row[key];
        });
        
        // Add the Google Maps link as a hyperlink in the 'GoogleMaps' column
        newRow['GoogleMaps'] = { f: `=HYPERLINK("${row['GoogleMaps']}","Go to Location")` };
        
        return newRow;
    }));
    // Append the sheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Addressing Data');
    
    // Save the workbook to a file with the custom file name
    XLSX.writeFile(wb, fileName);
}
