import express from 'express';
import cors from 'cors';
import XLSX from 'xlsx';

const app = express();
app.use(cors());
app.use(express.json());

app.post('/simular', (req, res) => {
  const { uf, pallets, meses } = req.body;

  try {
    const workbook = XLSX.readFile('simulacion.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    sheet['B2'].v = uf;
    sheet['B3'].v = pallets;
    sheet['B4'].v = meses;

    XLSX.writeFile(workbook, 'simulacion.xlsx');

    const resultado = {
      palletParking: uf * pallets * meses,
      tradicional: uf * pallets * meses * 1.2,
      ahorro: uf * pallets * meses * 0.8,
      tabla: Array.from({ length: meses }, (_, i) => ({
        mes: i + 1,
        entradas: Math.floor(Math.random() * 10),
        salidas: Math.floor(Math.random() * 10),
        stock: Math.floor(Math.random() * 50),
      })),
    };

    res.json(resultado);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(3000, () => console.log('Servidor en puerto 3000'));
