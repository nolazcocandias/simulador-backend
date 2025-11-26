import express from 'express';
import cors from 'cors';
import XLSX from 'xlsx';
import XLSX_CALC from 'xlsx-calc';

const app = express();
app.use(cors());
app.use(express.json());

function generarMovimientos(pallets, meses) {
  const inVals = new Array(meses).fill(0);
  const outVals = new Array(meses).fill(0);

  for (let i = 0; i < pallets; i++) {
    inVals[Math.floor(Math.random() * meses)]++;
  }

  let stock = 0;
  for (let i = 0; i < meses; i++) {
    stock += inVals[i];
    let maxOut = (i === meses - 1) ? stock : Math.floor(Math.random() * (stock + 1));
    outVals[i] = maxOut;
    stock -= maxOut;
  }
  return { inVals, outVals };
}

app.post('/simular', (req, res) => {
  try {
    let { uf, pallets, meses } = req.body;
    uf = Number(uf);
    pallets = Number(pallets);
    meses = Math.min(Math.max(Number(meses), 1), 12);

    if (!uf || !pallets || !meses) {
      return res.status(400).json({ error: 'Parámetros inválidos (uf, pallets, meses)' });
    }

    const workbook = XLSX.readFile('simulacion.xlsx');
    const sheetName = workbook.SheetNames.includes('cliente') ? 'cliente' : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
      return res.status(500).json({ error: 'Hoja no encontrada en Excel' });
    }

    const { inVals, outVals } = generarMovimientos(pallets, meses);

    for (let i = 0; i < 12; i++) {
      sheet['D' + (9 + i)] = { t: 'n', v: 0 };
      sheet['E' + (9 + i)] = { t: 'n', v: 0 };
    }
    for (let i = 0; i < meses; i++) {
      sheet['D' + (9 + i)] = { t: 'n', v: inVals[i] };
      sheet['E' + (9 + i)] = { t: 'n', v: outVals[i] };
    }

    // Escribir UF si corresponde
    sheet['W57'] = { t: 'n', v: uf };

    // Recalcular fórmulas con xlsx-calc
    XLSX_CALC(workbook);

    XLSX.writeFile(workbook, 'simulacion.xlsx');

    const tabla = [];
    for (let i = 0; i < meses; i++) {
      const r = 9 + i;
      const entradas = sheet['D' + r]?.v || 0;
      const salidas = sheet['E' + r]?.v || 0;
      const stock = sheet['G' + r]?.v || 0;
      tabla.push({ mes: i + 1, entradas, salidas, stock });
    }

    const palletParking = sheet['P103']?.v || 0;
    const tradicional = sheet['P104']?.v || 0;
    const ahorro = sheet['P105']?.v || (tradicional - palletParking);

    res.json({ palletParking, tradicional, ahorro, tabla });
  } catch (err) {
    console.error('Error en /simular:', err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));
