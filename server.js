const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const { Readable } = require('stream');
const archiver = require('archiver');
const app = express();
const port = 3000;

// Setup multer for file upload
const upload = multer({ storage: multer.memoryStorage() });

// Serve static files (HTML, CSS, JS)
app.use(express.static('public'));

// Helper function to clean and normalize brand names
function normalizeBrandName(name) {
    return name.trim().toLowerCase(); // Normalize brand names by trimming and lowercasing
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

    // All rows including the first row
    const rows = jsonData;

    // Organize data by brand with normalization
    const brands = {};

    rows.forEach(row => {
        const rawBrand = row[0];
        const normalizedBrand = normalizeBrandName(rawBrand);

        if (normalizedBrand) {
            if (!brands[normalizedBrand]) {
                brands[normalizedBrand] = [];
            }
            brands[normalizedBrand].push(row); // Add row to the brand's data
        }
    });

    // Filter brands to only include those with consistent names
    const uniqueBrands = {};
    const brandNames = Object.keys(brands);

    brandNames.forEach(brand => {
        const rows = brands[brand];
        const consistent = rows.every(row => normalizeBrandName(row[0]) === brand);
        if (consistent) {
            uniqueBrands[brand] = rows;
        }
    });

    const fileOption = req.body.option;
    const dateTimeString = getCurrentDateTimeString();

    if (fileOption === 'csv') {
        // Create CSV streams for each brand
        const fileStreams = Object.entries(uniqueBrands).map(([brand, rows]) => {
            // Exclude the header from the CSV content
            const csvContent = rows.map(row => row.join(",")).join("\n");
            const stream = new Readable();
            stream.push(csvContent);
            stream.push(null); // End of stream
            return {
                filename: `${brand}_${dateTimeString}.csv`,
                stream: stream
            };
        });

        // Generate ZIP file name with date and time
        const zipFileName = `files_${dateTimeString}.zip`;

        // Send the files as a zip
        const archive = archiver('zip', { zlib: { level: 9 } });

        res.setHeader('Content-disposition', `attachment; filename=${zipFileName}`);
        res.setHeader('Content-type', 'application/zip');

        archive.pipe(res);

        fileStreams.forEach(file => {
            archive.append(file.stream, { name: file.filename });
        });

        archive.finalize();
    } else if (fileOption === 'excel') {
        // Create a new workbook with sheets for each brand
        const newWorkbook = XLSX.utils.book_new();

        Object.entries(uniqueBrands).forEach(([brand, rows]) => {
            // Create a worksheet with rows only for each brand
            const ws = XLSX.utils.aoa_to_sheet(rows);
            XLSX.utils.book_append_sheet(newWorkbook, ws, brand);
        });

        // Generate Excel file name with date and time
        const excelFileName = `file_${dateTimeString}.xlsx`;

        const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'buffer' });

        res.setHeader('Content-disposition', `attachment; filename=${excelFileName}`);
        res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        res.end(excelBuffer);
    } else {
        res.status(400).send('Invalid option selected.');
    }
});

// Start server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
