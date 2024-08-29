const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const archiver = require('archiver');
const app = express();
const port = 3000;

// Setup multer for file upload
const upload = multer({ storage: multer.memoryStorage() });

// Serve static files (HTML, CSS, JS)
app.use(express.static('public'));

// Helper function to clean and normalize brand names
function normalizeBrandName(name) {
    return name ? name.trim().toLowerCase() : ''; // Handle possible undefined or null values
}

// Helper function to get current date and time as a string
function getCurrentDateTimeString() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0'); // Months are 0-based
    const day = String(now.getDate()).padStart(2, '0');
    const hour = String(now.getHours()).padStart(2, '0');
    const minute = String(now.getMinutes()).padStart(2, '0');
    const second = String(now.getSeconds()).padStart(2, '0');
    return `${year}-${month}-${day}_${hour}-${minute}-${second}`;
}

// Helper function to limit sheet name length to 31 characters
function limitSheetNameLength(name) {
    return name.substring(0, 31); // Limit the length to 31 characters
}

// Handle file upload and processing
app.post('/upload', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    const buffer = req.file.buffer;
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // All rows including the first row (headers)
    const rows = jsonData;

    // Organize data by brand with normalization
    const brands = {};

    rows.forEach((row, index) => {
        if (index === 0) return; // Skip the header row
        const rawBrand = row[1]; // Assuming the brand is in the second column
        const normalizedBrand = normalizeBrandName(rawBrand);

        if (normalizedBrand) {
            if (!brands[normalizedBrand]) {
                brands[normalizedBrand] = [rows[0]]; // Include headers initially
            }
            brands[normalizedBrand].push(row); // Add row to the brand's data
        }
    });

    // Filter brands to only include those with consistent names
    const uniqueBrands = {};
    const brandNames = Object.keys(brands);

    brandNames.forEach(brand => {
        const rows = brands[brand];
        const consistent = rows.slice(1).every(row => normalizeBrandName(row[1]) === brand); // Check consistency for all rows except header
        if (consistent) {
            uniqueBrands[brand] = rows;
        }
    });

    const dateTimeString = getCurrentDateTimeString();

    // Create ZIP archive
    const zip = archiver('zip', { zlib: { level: 9 } });
    res.setHeader('Content-disposition', `attachment; filename=files_${dateTimeString}.zip`);
    res.setHeader('Content-type', 'application/zip');

    zip.pipe(res);

    // Generate separate XLSX files for each brand and add them to the zip
    Object.entries(uniqueBrands).forEach(([brand, rows]) => {
        // Create a new workbook for each brand
        const newWorkbook = XLSX.utils.book_new();
        const limitedBrandName = limitSheetNameLength(brand);
        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(newWorkbook, ws, limitedBrandName);

        // Generate Excel file as a buffer
        const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'buffer' });
        const fileName = `${limitedBrandName}_${dateTimeString}.xlsx`;

        // Append the file buffer to the archive
        zip.append(excelBuffer, { name: fileName });
    });

    zip.finalize();
});

// Start server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
