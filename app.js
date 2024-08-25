const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');
const QRCode = require('qrcode');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.use(express.static('public'));

app.get('/', (req, res) => {
  res.render('index');
});

app.post('/submit', (req, res) => {
  const { name, sex, age, contact, email } = req.body;

  const filePath = './user_data.xlsx';

  let workbook;
  let worksheet;

  if (fs.existsSync(filePath)) {
    workbook = XLSX.readFile(filePath);
    worksheet = workbook.Sheets['Users'];
  } else {
    workbook = XLSX.utils.book_new();
    worksheet = XLSX.utils.json_to_sheet([]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Users');
  }

  XLSX.utils.sheet_add_json(
    worksheet,
    [{ Name: name, Sex: sex, Age: age, Contact: contact, Email: email }],
    { skipHeader: true, origin: -1 }
  );

  XLSX.writeFile(workbook, filePath);

  res.send('Form submitted successfully!');
});

app.get('/admin', (req, res) => {
  const qrLink = `${req.protocol}://${req.get('host')}`;
  QRCode.toDataURL(qrLink, (err, url) => {
    res.render('admin', { qrCode: url });
  });
});

app.get('/download', (req, res) => {
  const filePath = './user_data.xlsx';
  if (fs.existsSync(filePath)) {
    res.download(filePath);
  } else {
    res.send('No data available.');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
