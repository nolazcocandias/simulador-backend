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

function colLetterToNumber(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}
function numberToCol(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function ensureReferencedCellsExist(workbook) {
  const created = [];

  const cellRefRegex = /(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/g;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    // Iterate all cells in sheet looking for formulas
    for (const addr of Object.keys(sheet)) {
      if (!/^[A-Z]+[0-9]+$/.test(addr)) continue; // skip non-cell keys like '!ref'
      const cell = sheet[addr];
      if (cell && cell.f && typeof cell.f === 'string') {
        const formula = cell.f;
        let m;
        while ((m = cellRefRegex.exec(formula)) !== null) {
          const quotedSheet = m[1];
          const simpleSheet = m[2];
          const targetSheetName = quotedSheet || simpleSheet || sheetName;
          const startCol = m[3];
          const startRow = parseInt(m[4], 10);
          const endCol = m[5];
          const endRow = m[6] ? parseInt(m[6], 10) : null;

          const targetSheet = workbook.Sheets[targetSheetName];
          if (!workbook.Sheets[targetSheetName]) {
            // si la hoja objetivo no existe, no podemos crear celdas ahí; registramos y seguimos
            created.push({ sheet: targetSheetName, cell: null, note: 'sheet-not-found' });
            continue;
          }

          if (!endRow) {
            // celda simple, crear si no existe
            const ref = `${startCol}${startRow}`;
            if (!targetSheet[ref]) {
              targetSheet[ref] = { t: 'n', v: 0 };
              created.push({ sheet: targetSheetName, cell: ref });
            }
          } else {
            // rango: si es columna fija (A1:A10), expandimos filas
            if (startCol === endCol) {
              const col = startCol;
              const rStart = startRow;
              const rEnd = endRow;
              for (let r = rStart; r <= rEnd; r++) {
                const rr = `${col}${r}`;
                if (!targetSheet[rr]) {
                  targetSheet[rr] = { t: 'n', v: 0 };
                  created.push({ sheet: targetSheetName, cell: rr });
                }
              }
            } else {
              // rango multi-columna: creamos al menos las celdas de inicio y fin para evitar undefined
              const ref1 = `${startCol}${startRow}`;
              const ref2 = `${endCol}${endRow}`;
              if (!targetSheet[ref1]) {
                targetSheet[ref1] = { t: 'n', v: 0 };
                created.push({ sheet: targetSheetName, cell: ref1 });
              }
              if (!targetSheet[ref2]) {
                targetSheet[ref2] = { t: 'n', v: 0 };
                created.push({ sheet: targetSheetName, cell: ref2 });
              }
            }
          }
        }
      }
    }
  }

  return created;
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

    // Antes de ejecutar xlsx-calc: asegurar que las referencias usadas por las fórmulas existen como objetos de celda
    const createdRefs = ensureReferencedCellsExist(workbook);
    if (createdRefs.length > 0) {
      console.log('Se crearon celdas faltantes para evitar undefined al calcular:', createdRefs.slice(0, 200));
    } else {
      console.log('No se detectaron celdas faltantes a crear.');
    }

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
      const msg = String(calcErr && calcErr.message ? calcErr.message : calcErr);
      console.error('Error durante la ejecución de xlsx-calc:', msg);

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

      const definedNames = workbook.Workbook?.Names || workbook.Names || [];

      console.error('DebugCells:', JSON.stringify(debugCells, null, 2));
      console.error('Defined names:', JSON.stringify(definedNames, null, 2));

      return res.status(500).json({
        error: 'Error al ejecutar la función de cálculo',
        detail: msg,
        xlsxCalcMeta: meta,
        debugCells,
        definedNames,
        createdRefs: createdRefs.slice(0, 500)
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
