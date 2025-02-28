const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const app = express();
const path = require('path');

app.use(cors());
const upload = multer();
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// server.js (updated merge endpoint)
app.post('/merge', upload.fields([
  { name: 'firstSheet', maxCount: 1 },
  { name: 'secondSheet', maxCount: 1 }
]), (req, res) => {
  try {
      // Get user selections
      const firstMatch = req.body.firstMatch;
      const secondMatch = req.body.secondMatch;
      const columnsToAppend = JSON.parse(req.body.columnsToAppend);

      // Read Excel files
      const firstWB = xlsx.read(req.files.firstSheet[0].buffer, { type: 'buffer' });
      const secondWB = xlsx.read(req.files.secondSheet[0].buffer, { type: 'buffer' });

      // Convert to JSON
      const firstData = xlsx.utils.sheet_to_json(firstWB.Sheets[firstWB.SheetNames[0]]);
      const secondData = xlsx.utils.sheet_to_json(secondWB.Sheets[secondWB.SheetNames[0]]);

      // Create reference map
      const referenceMap = new Map();
      secondData.forEach(row => {
          const key = String(row[secondMatch])
              .replace(/\s+/g, '')
              .toLowerCase();
          referenceMap.set(key, columnsToAppend.reduce((acc, col) => {
              acc[col] = row[col];
              return acc;
          }, {}));
      });

      // Filter and merge data (CHANGED SECTION)
      const mergedData = firstData
          .filter(row => {
              const cleanKey = String(row[firstMatch])
                  .replace(/\s+/g, '')
                  .toLowerCase();
              return referenceMap.has(cleanKey);
          })
          .map(row => {
              const cleanKey = String(row[firstMatch])
                  .replace(/\s+/g, '')
                  .toLowerCase();
              return {
                  ...row,
                  ...referenceMap.get(cleanKey)
              };
          });

      // Create output workbook
      const ws = xlsx.utils.json_to_sheet(mergedData);
      const wb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(wb, ws, 'Merged Data');

      // Send response
      const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
      res.setHeader('Content-Disposition', 'attachment; filename="filtered_merged_data.xlsx"');
      res.contentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(buffer);

  } catch (error) {
      console.error('Merge error:', error);
      res.status(500).send(error.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
