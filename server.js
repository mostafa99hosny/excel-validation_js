const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));

// Utility functions and constants ported from Python
const EXPECTED_COLUMNS = [
  "asset_type", "asset_name", "final_value", "asset_usage_id", "value_base", "inspection_date", "production_capacity", "production_capacity_measuring_unit", "owner_name", "product_type", "market_approach", "market_approach_value", "cost_approach", "cost_approach_value", "country", "region", "city"
];
const MANDATORY_FIELDS = [
  "asset_type", "asset_name", "asset_usage_id", "value_base", "inspection_date", "final_value", "production_capacity", "production_capacity_measuring_unit", "owner_name", "product_type", "market_approach", "market_approach_value", "country", "region", "city"
];
const ASSET_USAGE_MIN = 38, ASSET_USAGE_MAX = 56;
const VALUE_BASE_MIN = 1, VALUE_BASE_MAX = 9;
const MARKET_APPROACH_ALLOWED = new Set([0, 1, 2]);

function isEmpty(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string' && value.trim() === '') return true;
  if (typeof value === 'string' && value.trim().toLowerCase() === 'n/a') return false;
  if (value === '0') return false;
  return false;
}

function toInt(value) {
  if (isEmpty(value)) return [false, null];
  let s = String(value).trim();
  if (s.toLowerCase().endsWith('.0')) s = s.slice(0, -2);
  if (s.includes('.')) return [false, null];
  const n = parseInt(s, 10);
  if (isNaN(n)) return [false, null];
  return [true, n];
}

function toFloat(value) {
  if (isEmpty(value)) return [false, null];
  const n = parseFloat(String(value).trim());
  if (isNaN(n)) return [false, null];
  return [true, n];
}

function appendMessage(existing, newMsg) {
  const s = String(existing).trim();
  if (!s || s.toLowerCase() === 'nan') return newMsg;
  return `${s} | ${newMsg}`;
}

// Validation logic (port of Python functions)
function checkMissingColumns(header) {
  return EXPECTED_COLUMNS.filter(c => !header.includes(c));
}

function validateFinalValueOnly(rows, header) {
  let issues = 0;
  const highlights = {};
  rows.forEach((row, idx) => {
    const colIdx = header.indexOf('final_value');
    if (colIdx === -1) return;
    let val = row[colIdx];
    let message = null;
    if (isEmpty(val)) {
      message = 'final_value is mandatory and cannot be empty';
    } else {
      const [ok, _intval] = toInt(val);
      if (!ok) message = 'Final value must be a non-decimal integer';
    }
    if (message) {
      issues++;
      highlights[`${idx},final_value`] = 'yellow';
      row[colIdx] = appendMessage(row[colIdx], message);
    }
  });
  return { rows, highlights, summary: [`Final Value issues: ${issues}`] };
}

function validateMandatoryOnly(rows, header) {
  let missingCount = 0;
  const highlights = {};
  rows.forEach((row, idx) => {
    MANDATORY_FIELDS.forEach(col => {
      const colIdx = header.indexOf(col);
      if (colIdx === -1) return;
      let val = row[colIdx];
      if (isEmpty(val)) {
        if (col === 'market_approach') return;
        if (col === 'market_approach_value') {
          const approachIdx = header.indexOf('market_approach');
          let approachRaw = approachIdx !== -1 ? row[approachIdx] : '';
          let approach = 0;
          if (!isEmpty(approachRaw)) {
            try { approach = parseInt(parseFloat(String(approachRaw).trim())); } catch { approach = null; }
          }
          if (approach === 0 || approach === null) return;
        }
        missingCount++;
        highlights[`${idx},${col}`] = 'yellow';
        row[colIdx] = appendMessage(row[colIdx], 'This mandatory field is empty');
      }
    });
  });
  return { rows, highlights, summary: [`Missing mandatory values: ${missingCount}`] };
}

function validateDatesOnly(rows, header) {
  let invalidCount = 0, autoFixed = 0;
  const highlights = {};
  const colIdx = header.indexOf('inspection_date');
  if (colIdx === -1) return { rows, highlights, summary: ["Column 'inspection_date' is missing"] };
  rows.forEach((row, idx) => {
    let val = row[colIdx];
    if (isEmpty(val)) {
      highlights[`${idx},inspection_date`] = 'yellow';
      row[colIdx] = appendMessage(row[colIdx], 'Date must be in dd-mm-YYYY format');
      invalidCount++;
      return;
    }
    // Try to parse date (simple dd-mm-yyyy or yyyy-mm-dd)
    let s = String(val).trim();
    let dt = null;
    let formatted = null;
    // Try dd-mm-yyyy
    let m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (m) {
      formatted = `${m[1]}-${m[2]}-${m[3]}`;
      dt = true;
    } else {
      // Try yyyy-mm-dd
      m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m) {
        formatted = `${m[3]}-${m[2]}-${m[1]}`;
        dt = true;
      }
    }
    if (dt && formatted) {
      row[colIdx] = formatted;
      autoFixed++;
    } else {
      highlights[`${idx},inspection_date`] = 'yellow';
      row[colIdx] = appendMessage(row[colIdx], 'Date must be in dd-mm-YYYY format');
      invalidCount++;
    }
  });
  let summary = [`Invalid dates: ${invalidCount}`];
  if (autoFixed) summary.push(`Dates auto-formatted: ${autoFixed}`);
  return { rows, highlights, summary };
}

function validateAll(rows, header) {
  let highlights = {};
  let summary = [];
  // 1) Mandatory non-empty
  let mand = validateMandatoryOnly(rows, header);
  highlights = { ...highlights, ...mand.highlights };
  summary = summary.concat(mand.summary);
  // 2) Final value integer
  let finalVal = validateFinalValueOnly(rows, header);
  highlights = { ...highlights, ...finalVal.highlights };
  summary = summary.concat(finalVal.summary);
  // 3) Dates
  let dates = validateDatesOnly(rows, header);
  highlights = { ...highlights, ...dates.highlights };
  summary = summary.concat(dates.summary);
  // 4) Additional numeric/range checks
  let extraIssues = 0;
  // asset_usage_id: integer 38..56
  let idxAssetUsage = header.indexOf('asset_usage_id');
  if (idxAssetUsage !== -1) {
    rows.forEach((row, idx) => {
      let val = row[idxAssetUsage];
      if (isEmpty(val)) return;
      const [ok, intval] = toInt(val);
      if (!ok || intval === null || intval < ASSET_USAGE_MIN || intval > ASSET_USAGE_MAX) {
        highlights[`${idx},asset_usage_id`] = 'yellow';
        row[idxAssetUsage] = appendMessage(row[idxAssetUsage], `asset_usage_id must be in [${ASSET_USAGE_MIN}-${ASSET_USAGE_MAX}]`);
        extraIssues++;
      }
    });
  }
  // value_base: integer 1..9
  let idxValueBase = header.indexOf('value_base');
  if (idxValueBase !== -1) {
    rows.forEach((row, idx) => {
      let val = row[idxValueBase];
      if (isEmpty(val)) return;
      const [ok, intval] = toInt(val);
      if (!ok || intval === null || intval < VALUE_BASE_MIN || intval > VALUE_BASE_MAX) {
        highlights[`${idx},value_base`] = 'yellow';
        row[idxValueBase] = appendMessage(row[idxValueBase], `value_base must be in [${VALUE_BASE_MIN}-${VALUE_BASE_MAX}]`);
        extraIssues++;
      }
    });
  }
  // market_approach: 0,1,2 (empty treated as 0)
  let idxMarketApproach = header.indexOf('market_approach');
  if (idxMarketApproach !== -1) {
    rows.forEach((row, idx) => {
      let val = row[idxMarketApproach];
      if (isEmpty(val)) return;
      const [ok, fval] = toFloat(val);
      if (!ok || fval === null || !MARKET_APPROACH_ALLOWED.has(parseInt(fval))) {
        highlights[`${idx},market_approach`] = 'yellow';
        row[idxMarketApproach] = appendMessage(row[idxMarketApproach], 'market_approach must be 0, 1, or 2');
        extraIssues++;
      }
    });
  }
  // market_approach_value: must be provided and numeric if approach in {1,2}; allowed empty if approach 0/empty
  let idxMarketApproachValue = header.indexOf('market_approach_value');
  if (idxMarketApproachValue !== -1) {
    rows.forEach((row, idx) => {
      let approachRaw = idxMarketApproach !== -1 ? row[idxMarketApproach] : '';
      let approachEmpty = isEmpty(approachRaw);
      let approach = 0;
      if (!approachEmpty) {
        try { approach = parseInt(parseFloat(String(approachRaw).trim())); } catch { approach = null; }
      }
      if (approach === 1 || approach === 2) {
        let val = row[idxMarketApproachValue];
        const [ok, fval] = toFloat(val);
        if (!ok || fval === null) {
          highlights[`${idx},market_approach_value`] = 'yellow';
          row[idxMarketApproachValue] = appendMessage(row[idxMarketApproachValue], 'Must be a number when approach is 1 or 2');
          extraIssues++;
        }
      }
    });
  }
  // production_capacity: if provided, must be non-negative number
  let idxProdCap = header.indexOf('production_capacity');
  if (idxProdCap !== -1) {
    rows.forEach((row, idx) => {
      let val = row[idxProdCap];
      if (isEmpty(val)) return;
      const [ok, fval] = toFloat(val);
      if (!ok || fval === null || fval < 0) {
        highlights[`${idx},production_capacity`] = 'yellow';
        row[idxProdCap] = appendMessage(row[idxProdCap], 'Must be a non-negative number');
        extraIssues++;
      }
    });
  }
  summary.push(`Additional rule violations: ${extraIssues}`);
  return { rows, highlights, summary };
}

// Serve main HTML page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Handle Excel file upload and validation
app.post('/validate', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];
    // Read header and rows
    const header = worksheet.getRow(1).values.slice(1); // ExcelJS is 1-based, values[0] is null
    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      rows.push(row.values.slice(1));
    });
    // Validation
    const missingCols = checkMissingColumns(header);
    if (missingCols.length > 0) {
      return res.status(400).send('The uploaded file is missing required columns: ' + missingCols.join(', '));
    }
    // Run all validations
    const { rows: validatedRows, highlights } = validateAll(rows, header);

    // Build a brand-new workbook to strip ALL formatting
    const outWb = new ExcelJS.Workbook();
    const outWs = outWb.addWorksheet(worksheet.name || 'Sheet1');

    // Helper: force black text; no background fill at all
    function applyBaseStyle(cell, fontSize = 11) {
      cell.fill = null; // no fill to prevent any background color
      cell.font = { name: 'Calibri', color: { argb: 'FF000000' }, size: fontSize };
      cell.border = undefined;
      cell.alignment = undefined;
      cell.numFmt = undefined;
    }

    // Write header as plain values (white background, black text)
    const headerRow = outWs.addRow(header);
    header.forEach((_, idx) => {
      const c = headerRow.getCell(idx + 1);
      applyBaseStyle(c, 11);
    });

    // أضف جميع الصفوف أولاً
    for (let i = 0; i < validatedRows.length; i++) {
      const rowArr = [];
      for (let j = 0; j < header.length; j++) {
        let value = validatedRows[i][j];
        let cellValue = value;
        // استرجع القيمة الأصلية فقط (قبل '|')
        if (typeof value === 'string' && value.includes('|')) {
          const [oldVal] = value.split('|');
          cellValue = oldVal.trim();
        }
        rowArr.push(cellValue);
      }
      outWs.addRow(rowArr);
    }
    // بعد إضافة جميع الصفوف، طبق التلوين والملاحظات بشكل صحيح
    for (let i = 0; i < validatedRows.length; i++) {
      const outRow = outWs.getRow(i + 2);
      for (let j = 0; j < header.length; j++) {
        const cell = outRow.getCell(j + 1);
        let value = validatedRows[i][j];
        // استخدم highlights فقط لتلوين الخلايا التي بها مشاكل
        if (highlights[`${i},${header[j]}`]) {
          cell.style = {};
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' }, // أصفر Excel
            bgColor: { argb: 'FFFFFF00' }
          };
          cell.font = { name: 'Arial', color: { argb: 'FF000000' }, bold: true, size: 13 };
          // أضف الرسالة كملاحظة
          if (typeof value === 'string' && value.includes('|')) {
            const [, ...msgParts] = value.split('|');
            const msg = msgParts.join('|').trim();
            if (msg) cell.note = msg;
          }
        } else {
          cell.style = {};
          cell.fill = undefined;
          cell.font = { name: 'Arial', color: { argb: 'FF218838' }, bold: true, size: 13 };
        }
      }
    }

    // تزيين رؤوس الأعمدة
    const headerRowDecorated = outWs.getRow(1);
    headerRowDecorated.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF21A366' } // أخضر مايكروسوفت
      };
      cell.font = { name: 'Cairo', color: { argb: 'FFFFFFFF' }, bold: true, size: 15 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = {
        top: { style: 'medium', color: { argb: 'FF21A366' } },
        left: { style: 'medium', color: { argb: 'FF21A366' } },
        bottom: { style: 'medium', color: { argb: 'FF21A366' } },
        right: { style: 'medium', color: { argb: 'FF21A366' } }
      };
    });
    outWs.views = [{ state: 'frozen', ySplit: 1 }];

    // Final sweep: ensure no fills remain and enforce black text
    outWs.eachRow((row) => {
      row.eachCell((cell) => {
        cell.fill = null;
        // Ensure text is black; if font exists, keep size, else set default 11
        const size = (cell.font && cell.font.size) ? cell.font.size : 11;
        cell.font = { name: 'Calibri', color: { argb: 'FF000000' }, size };
      });
    });

    // Save and send file
    const outPath = filePath + '_validated.xlsx';
    await outWb.xlsx.writeFile(outPath);
    res.download(outPath, 'validated.xlsx', () => {
      fs.unlinkSync(filePath);
      fs.unlinkSync(outPath);
    });
  } catch (err) {
    res.status(500).send('Error processing file: ' + err.message);
  }
});

// Add endpoint to return summary and preview data as JSON
app.post('/validate-preview', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];
    const header = worksheet.getRow(1).values.slice(1);
    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      rows.push(row.values.slice(1));
    });
    const missingCols = checkMissingColumns(header);
    if (missingCols.length > 0) {
      return res.status(400).json({ error: 'The uploaded file is missing required columns: ' + missingCols.join(', ') });
    }
    let validated, total = null;
    const type = req.query.type;
    if (type === 'date') {
      validated = validateDatesOnly(rows, header);
    } else if (type === 'mandatory') {
      validated = validateMandatoryOnly(rows, header);
    } else if (type === 'final') {
      validated = validateFinalValueOnly(rows, header);
    } else if (type === 'sum') {
      // Only sum final_value
      let idx = header.indexOf('final_value');
      if (idx === -1) return res.status(400).json({ error: 'No final_value column' });
      total = rows.reduce((acc, row) => {
        let v = row[idx];
        // Remove appended messages if present
        if (typeof v === 'string') {
          v = v.split('|')[0];
          v = v.trim();
          v = v.replace(/,/g, '');
        }
        // If empty string or null, treat as 0
        if (v === '' || v === null || v === undefined) v = 0;
        let num = parseFloat(v);
        if (isNaN(num)) num = 0;
        return acc + num;
      }, 0);
      fs.unlinkSync(filePath);
      return res.json({ total });
    } else {
      validated = validateAll(rows, header);
    }
    const { rows: validatedRows, highlights, summary } = validated;
    const preview = validatedRows.map((row, i) => {
      const obj = {};
      header.forEach((col, j) => {
        obj[col] = {
          value: row[j],
          highlight: !!highlights[`${i},${col}`]
        };
      });
      return obj;
    });
    fs.unlinkSync(filePath);
    res.json({ header, preview, summary });
  } catch (err) {
    res.status(500).json({ error: 'Error processing file: ' + err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
