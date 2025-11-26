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
  // Distribuye los pallets aleatoriamente entre meses (suma inVals = pallets)
  const inVals = new Array(meses).fill(0);
  for (let i = 0; i < pallets; i++) {
    inVals[Math.floor(Math.random() * meses)]++;
  }

  // Genera salidas (outVals) asegurando que no se supere el stock en cada mes
  const outVals = new Array(meses).fill(0);
  let stock = 0;
  for (let i = 0; i < meses; i++) {
    stock += inVals[i];
    if (i === meses - 1) {
      // último mes vaciamos todo
      outVals[i] = stock;
    } else {
      // sacamos entre 0 y stock (inclusive), de forma aleatoria
      outVals[i] = Math.floor(Math.random() * (stock + 1));
    }
    stock -= outVals[i];
  }

  // Como comprobación, stock debe ser 0 al final
  if (stock !== 0) {
    // ajuste por si hubiera algún desfase numérico (no debería ocurrir)
    outVals[meses - 1] += stock;
    stock = 0;
  }

  return { inVals, outVals };
}

async function loadXlsxCalc() {
  try {
    // Import dinámico para permitir compatibilidades ESM/CJS
    const mod = await import('xlsx-calc');
    const keys = Object.keys(mod);
    console.log('xlsx-calc import keys:', keys);

    // Posibles formas de obtener la función de cálculo:
    if (typeof mod === 'function') return { fn: mod, meta: keys };
    if (mod.default && typeof mod.default === 'function') return { fn: mod.default, meta: keys };
    if (mod.calc && typeof mod.calc === 'function') return { fn: mod.calc, meta: keys };

    // Si ninguno es función, devolvemos las keys para diagnóstico
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

    // Leer con opciones razonables para conservar fórmulas/metadatos
    const workbook = XLSX.readFile(excelPath, { cellNF: true, cellDates: true });
    const sheetName = workbook.SheetNames.includes('cliente') ? 'cliente' : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
      return res.status(500).json({ error: 'Hoja no encontrada en Excel', sheetName, sheetNames: workbook.SheetNames });
    }

    const { inVals, outVals } = generarMovimientos(pallets, meses);

    // Reescribir D9:D20 y E9:E20 (12 meses) — primero limpiar, luego setear los meses activos
    for (let i = 0; i < 12; i++) {
      sheet['D' + (9 + i)] = { t: 'n', v: 0 };
      sheet['E' + (9 + i)] = { t: 'n', v: 0 };
    }
    for (let i = 0; i < meses; i++) {
      sheet['D' + (9 + i)] = { t: 'n', v: inVals[i] };
      sheet['E' + (9 + i)] = { t: 'n', v: outVals[i] };
    }

    // Escribir UF si corresponde (W57 en tu excel según servidor actual)
    sheet['W57'] = { t: 'n', v: uf };

    // Cargar y detectar cómo invocar xlsx-calc
    const { fn: calcFn, meta } = await loadXlsxCalc();
    if (!calcFn) {
      console.error('No se detectó función de cálculo en xlsx-calc. meta:', meta);
      return res.status(500).json({ error: 'xlsx-calc no exporta función invocable', meta });
    }

    // Intentar ejecutar la función de cálculo y capturar errores con contexto
    try {
      console.log('Invocando función de cálculo (xlsx-calc)...');
      // La mayoría de las versiones usan: calcFn(workbook)
      await calcFn(workbook);
      console.log('Cálculo completado satisfactoriamente.');
    } catch (calcErr) {
      console.error('Error durante la ejecución de xlsx-calc:', calcErr && calcErr.stack || calcErr);
      // Si el error contiene referencia a una celda, incluimos eso para diagnóstico
      return res.status(500).json({
        error: 'Error al ejecutar la función de cálculo',
        detail: String(calcErr && calcErr.message || calcErr),
        xlsxCalcMeta: meta
      });
    }

    // Opcional: escribir el workbook de vuelta (si quieres persistir)
    try {
      XLSX.writeFile(workbook, excelPath);
    } catch (writeErr) {
      console.warn('No fue posible sobrescribir el archivo excel (esto puede estar bien en entornos readonly):', writeErr && writeErr.message);
    }

    // Extraer la tabla y KPIs
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
