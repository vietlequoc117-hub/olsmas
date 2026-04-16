import express from 'express';
import { createServer as createViteServer } from 'vite';
import multer from 'multer';
import ExcelJS from 'exceljs';
import cors from 'cors';
import path from 'path';
import fs from 'fs';

const app = express();
const port = 3000;

app.use(cors());
app.use(express.json());

const upload = multer({ storage: multer.memoryStorage() });

// Helper to safely get string value from a cell
const getCellString = (cell: any): string => {
  if (!cell) return '';
  try {
    const val = cell.value;
    if (val === null || val === undefined) return '';
    
    // Handle Rich Text
    if (val && typeof val === 'object' && 'richText' in val) {
      return val.richText.map((rt: any) => rt.text || '').join('');
    }
    
    // Handle Formulas
    if (val && typeof val === 'object' && 'result' in val) {
      return String(val.result ?? '');
    }

    // Handle other objects (like Hyperlinks)
    if (val && typeof val === 'object' && 'text' in val) {
      return String(val.text ?? '');
    }

    return String(val);
  } catch (e) {
    return '';
  }
};

// Helper to extract class code (e.g., "10A1" from "Lớp 10A1" or "Tin học 10A1")
const extractClassCode = (text: string): string => {
  const clean = text.trim().toUpperCase();
  const match = clean.match(/(\d{1,2}[A-Z]\d{0,2})/);
  return match ? match[1].toLowerCase() : clean.toLowerCase().replace(/\s+/g, '');
};

// Helper to clean and normalize text
const cleanText = (text: any): string => {
  if (text === null || text === undefined) return '';
  try {
    // Chuyển về chuỗi an toàn, xóa khoảng trắng và chuẩn hóa
    const str = String(text);
    return str.trim().toLowerCase().replace(/\s+/g, ' ');
  } catch (e) {
    return '';
  }
};

// API route for processing
app.post('/api/sync-grades', upload.fields([
  { name: 'sources', maxCount: 3 },
  { name: 'target', maxCount: 1 }
]), async (req: any, res) => {
  try {
    const sourceFiles = req.files['sources'];
    const targetFile = req.files['target']?.[0];

    if (!sourceFiles || sourceFiles.length === 0 || !targetFile) {
      return res.status(400).json({ error: 'Thiếu file nguồn hoặc file đích.' });
    }

    const masterData = new Map<string, any>();
    const stats = {
      totalSourceStudents: 0,
      classStats: {} as Record<string, { count: number, totalScore: number, min: number, max: number }>
    };

    // 1. Process Source Files
    for (const file of sourceFiles) {
      try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(file.buffer);

        let worksheet = workbook.worksheets.find(s => s.name.includes('Danh sách HS'));

        if (!worksheet) {
          console.warn(`No worksheet named 'Danh sách HS' found in ${file.originalname}`);
          continue;
        }

        console.log(`Processing source sheet: ${worksheet.name} from ${file.originalname}`);

        let startRow = -1;
        let sourceNameCol = 2;
        let sourceClassCol = 4;
        let sourceScoreCol = 10;

        worksheet.eachRow((row, rowNumber) => {
          if (startRow !== -1) return;
          
          row.eachCell((cell, colNumber) => {
            const val = cleanText(getCellString(cell));
            if (val.includes('họ và tên') || val === 'họ tên') sourceNameCol = colNumber;
            if (val === 'lớp') sourceClassCol = colNumber;
            if (val.includes('tin học')) sourceScoreCol = colNumber;
          });

          const cell = row.getCell(1);
          const stt = getCellString(cell);
          if (cleanText(stt) === '1') {
            startRow = rowNumber;
          }
        });

        if (startRow !== -1) {
          worksheet.eachRow((row, rowNumber) => {
            if (rowNumber < startRow) return;
            
            const name = cleanText(getCellString(row.getCell(sourceNameCol)));
            const classRaw = getCellString(row.getCell(sourceClassCol));
            const className = extractClassCode(classRaw);
            const scoreVal = row.getCell(sourceScoreCol).value;
            const score = typeof scoreVal === 'number' ? scoreVal : parseFloat(String(scoreVal));

            if (name && className) {
              const key = `${name}|${className}`;
              masterData.set(key, score);
              stats.totalSourceStudents++;
              
              if (!stats.classStats[className]) {
                stats.classStats[className] = { count: 0, totalScore: 0, min: Infinity, max: -Infinity };
              }
              if (!isNaN(score)) {
                stats.classStats[className].count++;
                stats.classStats[className].totalScore += score;
                stats.classStats[className].min = Math.min(stats.classStats[className].min, score);
                stats.classStats[className].max = Math.max(stats.classStats[className].max, score);
              }
            }
          });
        }
      } catch (err) {
        console.error(`Error processing source file ${file.originalname}:`, err);
      }
    }

    // 2. Process Target File
    const targetWorkbook = new ExcelJS.Workbook();
    try {
      await targetWorkbook.xlsx.load(targetFile.buffer);
    } catch (err) {
      console.error('Error loading target workbook:', err);
      throw new Error('Không thể đọc file đích. Vui lòng kiểm tra định dạng file.');
    }

    let updatedCount = 0;
    let totalStudentsInTarget = 0;

    for (const sheet of targetWorkbook.worksheets) {
      const sheetName = sheet.name;
      const classFromSheet = extractClassCode(sheetName);

      let startRow = 14; // Bắt đầu từ dòng 14 như yêu cầu
      let colScore = 8; // Cột H (ĐĐG GK)
      let targetNameCol = 4; // Cột D (Họ và tên) theo hình ảnh

      sheet.eachRow((row, rowNumber) => {
        if (rowNumber > 20) return;
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const cellValue = cleanText(getCellString(cell));
          if (cellValue.includes('họ và tên') || cellValue === 'họ tên') targetNameCol = colNumber;
          if (cellValue.includes('đđg gk') || cellValue === 'gk') colScore = colNumber;
        });
      });

      if (startRow !== -1) {
        sheet.eachRow((row, rowNumber) => {
          if (rowNumber < startRow) return;
          const nameVal = getCellString(row.getCell(targetNameCol));
          if (!nameVal || cleanText(nameVal) === '') return;
          
          totalStudentsInTarget++;
          const nameClean = cleanText(nameVal);
          const key = `${nameClean}|${classFromSheet}`;

          if (masterData.has(key)) {
            const score = masterData.get(key);
            const scoreCell = row.getCell(colScore);
            if (scoreCell) {
              scoreCell.value = score;
              updatedCount++;
            }
          }
        });
      }
    }

    // 3. Send back the processed file
    let buffer;
    try {
      buffer = await targetWorkbook.xlsx.writeBuffer();
    } catch (err) {
      console.error('Error writing workbook to buffer:', err);
      throw new Error('Lỗi khi tạo file kết quả.');
    }
    
    const aiSummary = {
      updatedCount,
      totalStudentsInTarget,
      totalSourceStudents: stats.totalSourceStudents,
      classes: Object.entries(stats.classStats).map(([name, s]) => ({
        name,
        count: s.count,
        avg: s.count > 0 ? (s.totalScore / s.count).toFixed(2) : 0,
        min: s.min === Infinity ? 0 : s.min,
        max: s.max === -Infinity ? 0 : s.max
      }))
    };

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=BangDiem_KetQua_DongBo.xlsx');
    res.setHeader('X-Updated-Count', String(updatedCount));
    res.setHeader('X-AI-Summary', JSON.stringify(aiSummary));
    res.send(buffer);

  } catch (error: any) {
    console.error('Error processing Excel:', error);
    res.status(500).json({ error: 'Lỗi xử lý file: ' + error.message });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(port, '0.0.0.0', () => {
    console.log(`Server running at http://localhost:${port}`);
  });
}

startServer();
