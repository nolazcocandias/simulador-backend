import express from 'express';
import cors from 'cors';
import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());

function generarMovimientos(pallets, meses) {
  const inVals = new Array(meses).fill(0);
  for (let i = 0; i < pallets; i++) {
    inVals[Math.floor(Math.random() * meses)]++;
  }

  const outVals = new Array(meses).fill(0);
  let stock = 0;
  for (let i = 0; i < meses; i++) {
    stock += inVals[i];
    if (i === meses - 1) {
      outVals[i] = stock;
    } else {
      outVals[i] = Math.floor(Math.random() * (stock + 1));
    }
    stock -= outVals[i];
  }

  if (stock !== 0) {
    outVals[meses - 1] += stock;
    stock = 0;
  }

  return { inVals, outVals };
}

async function loadXlsxCalc() {
  try {
    const mod = await import('xlsx-calc');
    const keys = Object.keys(mod);
    console.log('xlsx-calc import keys:', keys);

    if (typeof mod === 'function') return { fn: mod, meta: keys };
    if (mod.default && typeof mod.default === 'function') return { fn: mod.default, meta: keys };
    if (mod.calc && typeof mod.calc === 'function') return { fn: mod.calc, meta: keys };

    return { fn: null, meta: keys };
  } catch (err) {
    console.error('Error importando xlsx-calc:', err);
    return { fn: null, meta: ['import-error', String(err.message || err)] };
  }
}

app.post('/simular', async (req, res) => {
  try {
    let { uf, pallets, meses } = req.body;
    uf = Number(uf);
    pallets = Number(pallets);
    meses = Math.min(Math.max(Number(meses), 1), 12);

    if (!uf || !pallets || !meses) {
      return res.status(400).json({ error: 'Parámetros inválidos (uf, pallets, meses)' });
    }

    const excelPath = path.resolve(__dirname, 'simulacion.xlsx');
    console.log('Intentando leer Excel en:', excelPath);
    if (!fs.existsSync(excelPath)) {
      console.error('Archivo Excel NO encontrado en la ruta indicada:', excelPath);
      return res.status(500).json({ error: 'Archivo Excel no encontrado en servidor', path: excelPath });
    }

    const workbook = XLSX.readFile(excelPath, { cellNF: true, cellDates: true });
    const sheetName = workbook.SheetNames.includes('cliente') ? 'cliente' : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
      return res.status(500).json({ error: 'Hoja no encontrada en Excel', sheetName, sheetNames: workbook.SheetNames });
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

    sheet['W57'] = { t: 'n', v: uf };

    const { fn: calcFn, meta } = await loadXlsxCalc();
    if (!calcFn) {
      console.error('No se detectó función de cálculo en xlsx-calc. meta:', meta);
      return res.status(500).json({ error: 'xlsx-calc no exporta función invocable', meta });
    }

    try {
      console.log('Invocando función de cálculo (xlsx-calc)...');
      await calcFn(workbook);
      console.log('Cálculo completado satisfactoriamente.');
    } catch (calcErr) {
      // Debug avanzado: extraer celdas mencionadas en el mensaje de error y devolver su estado
      const msg = String(calcErr && calcErr.message ? calcErr.message : calcErr);
      console.error('Error durante la ejecución de xlsx-calc:', msg);

      // Extraer referencias tipo A1 (p.ej. D103, W57, J96)
      const matches = (msg.match(/([A-Z]+\d+)/g) || []).map(s => s.trim());
      const uniqueCells = [...new Set(matches)];

      const debugCells = uniqueCells.map(cellRef => {
        const obj = sheet[cellRef];
        return {
          cell: cellRef,
          present: !!obj,
          f: obj?.f ?? null,
          v: obj?.v ?? null,
          t: obj?.t ?? null
        };
      });

      // Nombres definidos en el workbook (si existen)
      const definedNames = workbook.Workbook?.Names || workbook.Names || [];

      // Loguear para Render
      console.error('DebugCells:', JSON.stringify(debugCells, null, 2));
      console.error('Defined names:', JSON.stringify(definedNames, null, 2));

      return res.status(500).json({
        error: 'Error al ejecutar la función de cálculo',
        detail: msg,
        xlsxCalcMeta: meta,
        debugCells,
        definedNames
      });
    }

    try {
      XLSX.writeFile(workbook, excelPath);
    } catch (writeErr) {
      console.warn('No fue posible sobrescribir el archivo excel (puede ser readonly):', writeErr && writeErr.message);
    }

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
    const ahorro = sheet['P105']?.v ?? (tradicional - palletParking);

    res.json({ palletParking, tradicional, ahorro, tabla });
  } catch (err) {
    console.error('Error en /simular (catch general):', err && err.stack || err);
    res.status(500).json({ error: 'Error inesperado en servidor', detail: String(err && err.message || err) });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));
