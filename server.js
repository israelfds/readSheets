const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const app = express();
const port = 3000;

// Configurar o middleware Multer para lidar com uploads de arquivos
const storage = multer.memoryStorage(); // Armazenar o arquivo em memória
const upload = multer({ storage });

// Pasta onde você deseja salvar o JSON
const jsonFolder = path.join(__dirname, 'upJson');

// Certificar-se de que a pasta existe
if (!fs.existsSync(jsonFolder)) {
  fs.mkdirSync(jsonFolder);
}

app.post('/api/uploadSheets', upload.single('excelFile'), async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: 'Nenhum arquivo foi enviado.' });
      }
  
      const buffer = req.file.buffer;
  
      // Ler o arquivo Excel
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
  
      // Processar o arquivo Excel e criar um JSON
      const jsonResult = {};
  
      let lastWorksheet;
      let headers; // Defina a variável headers fora do loop
  
      workbook.eachSheet((worksheet) => {
        const sheetData = [];
  
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) {
            // Cabeçalhos da planilha
            headers = row.values;
            return;
          }
  
          const rowData = row.values;
          const rowObject = {};
  
          for (let i = 0; i < headers.length; i++) {
            rowObject[headers[i]] = rowData[i];
          }
  
          sheetData.push(rowObject);
        });
  
        jsonResult[worksheet.name] = sheetData;
  
        lastWorksheet = worksheet;
      });
  
      if (lastWorksheet) {
        const jsonFileName = `${lastWorksheet.name}.json`;
        const jsonFilePath = path.join(jsonFolder, jsonFileName);
        fs.writeFileSync(jsonFilePath, JSON.stringify(jsonResult, null, 2));
      }
  
      res.json({ result: jsonResult });
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: 'Ocorreu um erro ao processar o arquivo Excel.' });
    }
  });
  

app.listen(port, () => {
  console.log(`Servidor está rodando na porta ${port}`);
});
