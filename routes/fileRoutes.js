const express = require('express');
const fs = require('fs').promises;
const path = require('path');
const xlsx = require('xlsx');

const router = express.Router();
const ALLOWED_EXTENSION = '.xlsx';

// Helper function to get the absolute file path
const getFilePath = (fileName) => path.join(__dirname, '..', 'storage', path.basename(fileName));

// Validate that the file has the correct extension
const isValidFileType = (fileName) => path.extname(fileName).toLowerCase() === ALLOWED_EXTENSION;

// Middleware to enforce file type validation
const validateFileExtension = (req, res, next) => {
  const fileName = req.query.fileName || req.body.fileName || req.body.oldName || req.body.newName;
  if (!fileName || !isValidFileType(fileName)) {
    return res.status(400).json({ error: `Only ${ALLOWED_EXTENSION} files are allowed.` });
  }
  next();
};

// READ EXCEL FILE
router.get('/read', validateFileExtension, async (req, res) => {
  try {
    const filePath = getFilePath(req.query.fileName);
    const fileBuffer = await fs.readFile(filePath);
    
    // Parse Excel file
    const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
    
    // Convert the first sheet to JSON
    const sheetName = workbook.SheetNames[0];
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    res.json({ fileName: req.query.fileName, data: sheetData });
  } catch (err) {
    res.status(500).json({ error: 'Error reading the Excel file.' });
  }
});

// WRITE EXCEL FILE
router.post('/write', validateFileExtension, async (req, res) => {
  try {
    const { fileName, content } = req.body;
    const filePath = getFilePath(fileName);

    // Convert JSON content to worksheet
    const worksheet = xlsx.utils.json_to_sheet(content);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Write the file
    await fs.writeFile(filePath, xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' }));

    res.json({ message: 'File written successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// APPEND TO EXCEL FILE
router.post('/append', validateFileExtension, async (req, res) => {
  try {
    const { fileName, content } = req.body;
    const filePath = getFilePath(fileName);

    // Read existing workbook
    const fileBuffer = await fs.readFile(filePath);
    const workbook = xlsx.read(fileBuffer, { type: 'buffer' });

    // Convert content to worksheet and append
    const sheetName = workbook.SheetNames[0];
    const existingData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    const updatedData = [...existingData, ...content];

    const worksheet = xlsx.utils.json_to_sheet(updatedData);
    workbook.Sheets[sheetName] = worksheet;

    await fs.writeFile(filePath, xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' }));

    res.json({ message: 'Content appended successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// DELETE FILE
router.delete('/delete', validateFileExtension, async (req, res) => {
  try {
    await fs.unlink(getFilePath(req.query.fileName));
    res.json({ message: 'File deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.put('/rename', validateFileExtension, async (req, res) => {
  const { oldName, newName } = req.body;

  if (!oldName || !newName) {
    return res.status(400).json({ error: 'Both old and new file names are required' });
  }

  // Ensure both names have the correct .xlsx extension
  if (!isValidFileType(oldName) || !isValidFileType(newName)) {
    return res.status(400).json({ error: 'Only .xlsx files are allowed for renaming.' });
  }

  const oldFilePath = getFilePath(oldName);
  const newFilePath = getFilePath(newName);

  try {
    // Check if the old file exists
    await fs.access(oldFilePath);

    // Rename the file
    await fs.rename(oldFilePath, newFilePath);
    res.json({ message: 'File renamed successfully' });
  } catch (err) {
    res.status(500).json({ error: `Error renaming file: ${err.message}` });
  }
});

module.exports = router;
