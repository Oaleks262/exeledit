const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { google } = require('googleapis');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Налаштування Google API
const credentials = require('./credentials.json');
const spreadsheetId = '1Rdy0e3ZUbfz9Ufszfj6kCh9iwVtgF9Yi2bhustXI1x8'; // Замініть на ваш ID таблиці

const auth = new google.auth.GoogleAuth({
  credentials: credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
});

const sheets = google.sheets('v4');

// Функція нормалізації тексту
const normalize = (str) => str?.trim().toLowerCase().replace(/\s+/g, '');

// Віддача статичних файлів із директорії public
app.use(express.static('public'));

// Маршрут для головної сторінки (index.html)
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Роут для завантаження файлів
app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    // Перевірка типу файлу
    const fileType = req.file.mimetype;
    if (fileType !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      return res.status(400).send('Invalid file type. Please upload an Excel file.');
    }

    // Читання ексель файлу
    const filePath = path.join(__dirname, req.file.path);
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const products = xlsx.utils.sheet_to_json(sheet);
    console.log('Products from uploaded file:', products);

    // Підключення до Google Sheets
    const authClient = await auth.getClient();
    const response = await sheets.spreadsheets.values.get({
      auth: authClient,
      spreadsheetId,
      range: 'Товари з кодами!A:B', // Перевірте правильність діапазону
    });

    const googleData = response.data.values;
    const productCodes = {};

    googleData.forEach(([name, code]) => {
      productCodes[normalize(name)] = code;
    });

    console.log('Google Sheets data:', productCodes);

    // Оновлення ексель файлу з кодами товарів
    products.forEach((product) => {
      const productName = normalize(product['Name'] || product['A']); // Використовуйте правильний ключ
      if (productCodes[productName]) {
        product['Code'] = productCodes[productName];
      } else {
        product['Code'] = 'Not Found';
      }
    });

    console.log('Updated products:', products);

    // Створення нового ексель файлу з кодами
    const newSheet = xlsx.utils.json_to_sheet(products, { header: ['Name', 'Code'] });
    const newWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'UpdatedProducts');

    const newFilePath = path.join(__dirname, 'updated_products.xlsx');
    xlsx.writeFile(newWorkbook, newFilePath);
    console.log('New file created:', newFilePath);

    // Відправка файлу назад клієнту
    res.download(newFilePath, 'updated_products.xlsx', () => {
      fs.unlinkSync(filePath); // Видалення тимчасового файлу
      fs.unlinkSync(newFilePath); // Видалення оновленого файлу
    });
  } catch (error) {
    console.error('Error processing the file:', error);
    res.status(500).send('Error processing the file');
  }
});

// Запуск сервера
app.listen(4444, () => {
  console.log('Server running on port 4444');
});
