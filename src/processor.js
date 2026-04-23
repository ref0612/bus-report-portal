'use strict';
const Papa = require('papaparse');
const XLSX = require('xlsx');

// ─── CSV parser ───────────────────────────────────────────────────────────────
// Column layout (19 cols, 0-indexed):
//  [0]  Numero de Pasaje Operador  → _pb
//  [2]  Emitido en  (DD/MM/YYYY HH:MM AM/PM)  → _date  ← EMISSION date
//  [3]  Fecha de viaje  → travel date (NOT used for grouping)
//  [4]  Origen  → _origin
//  [5]  Destino → _dest
//  [6]  Asientos → _seat
//  [7]  Nombre del operador → _operator
//  [9]  Nombre de Cliente → _name
//  [10] Correo → _email
//  [12] Tipo de reserva → _channel
//  [13] PG Type → _gateway
//  [14] PG StatusBrowser/Version → _pgStatus
//  [16] Booking Source (mostly '-'/blank — not useful)
//  [17] Platform/site (pasajebus, cormarbus.cl, etc.) → _platform
//  [18] Precio del pasaje → _price
function parseCSVDate(str) {
  // Accepts "DD/MM/YYYY HH:MM AM/PM" or "DD/MM/YYYY HH:MM"
  if (!str) return null;
  const m = String(str).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? `${m[3]}-${m[2]}-${m[1]}` : null;
}

function cleanStr(v) {
  const s = String(v || '').trim();
  return (s && s !== '-') ? s : null;
}

function parseCSV(buffer) {
  const text = buffer.toString('utf-8');
  const raw  = Papa.parse(text, { header: false, skipEmptyLines: true });
  if (raw.data.length < 2) return [];
  return raw.data.slice(1)
    .filter(r => r[0] && String(r[0]).startsWith('PB'))
    .filter(r => !String(r[17] || '').toLowerCase().includes('total'))
    .map(r => ({
      _pb:       r[0]  || null,
      _date:     parseCSVDate(r[2]),   // col[2] = "Emitido en" (emission date)
      _origin:   r[4]  || null,
      _dest:     r[5]  || null,
      _seat:     r[6]  || null,
      _operator: r[7]  || null,
      _name:     r[9]  || null,
      _email:    r[10] || null,
      _channel:  cleanStr(r[12]),
      _gateway:  cleanStr(r[13]),
      _pgStatus: r[14] || null,
      _platform: cleanStr(r[17]),      // col[17] = platform/site (pasajebus, cormarbus.cl)
      _price:    parseFloat(r[18]) || 0,
    }));
}

// ─── KONNECT sales Excel ───────────────────────────────────────────────────────
// Tries 3 strategies in order:
//  1. Any sheet with Date in col[0] + number in col[1] → pivot table (Hoja1 style)
//  2. 'Datos' sheet → raw transaction data, group by Fecha (col 13), exclude DEVUELTO
//  3. Any sheet with raw data containing Fecha-like Date column → same approach
function parseSalesKonnect(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });

  // Strategy 1: find pivot sheet (Date | count pairs)
  for (const name of wb.SheetNames) {
    const ws   = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const dateRows = rows.filter(r => r[0] instanceof Date && typeof r[1] === 'number' && r[1] > 0);
    if (dateRows.length >= 3) {
      // looks like a pivot — use it
      return dateRows.map(r => {
        const d = r[0];
        return {
          date:  `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`,
          count: Number(r[1]) || 0,
        };
      });
    }
  }

  // Strategy 2 & 3: raw transactions sheet
  for (const name of wb.SheetNames) {
    const ws   = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (rows.length < 10) continue;

    // find which column has the most Date objects (that's the Fecha column)
    const header  = rows[0];
    const sample  = rows.slice(1, 20);
    let dateCol   = -1;
    let descCol   = -1;
    let maxDates  = 0;

    for (let c = 0; c < (header.length || 30); c++) {
      const cnt = sample.filter(r => r[c] instanceof Date).length;
      if (cnt > maxDates) { maxDates = cnt; dateCol = c; }
    }
    // find description column (contains ticket codes like TS...)
    for (let c = 0; c < header.length; c++) {
      const vals = sample.map(r => String(r[c] || ''));
      if (vals.some(v => v.startsWith('TS') || v.includes('DEVUELTO'))) { descCol = c; break; }
    }
    if (dateCol < 0 || maxDates < 5) continue;

    const dayMap = {};
    rows.slice(1).forEach(r => {
      if (!(r[dateCol] instanceof Date)) return;
      if (descCol >= 0 && String(r[descCol] || '').includes('DEVUELTO')) return;
      const d   = r[dateCol];
      const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
      dayMap[key] = (dayMap[key] || 0) + 1;
    });

    const result = Object.entries(dayMap)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([date, count]) => ({ date, count }));

    if (result.length >= 3) return result;
  }

  return []; // no usable data found
}

// ─── API sales Excel ───────────────────────────────────────────────────────────
function parseSalesAPI(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  if (rows.length < 2) return { byDay: [], tickets: [] };

  // Positional indices (confirmed from file inspection)
  const iEmitido  = 2;
  const iCancel   = 3;
  const iReserva  = 9;
  const iOperador = 0;
  const iOrigen   = 4;
  const iDestino  = 5;
  const iPrecio   = 12;
  const iPrecioAPI= 13;
  const iPlatform = 25;
  const iPGType   = 26;
  const iMovil    = 27;
  const iCliente  = 6;
  const iCorreo   = 8;

  function parseDate(str) {
    if (!str) return null;
    const m = String(str).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    return m ? `${m[3]}-${m[2]}-${m[1]}` : null;
  }
  function parsePrice(v) {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    return parseFloat(String(v).replace(/\./g, '').replace(',', '.')) || 0;
  }
  function isCancelled(v) {
    return v !== null && v !== undefined && String(v).trim().length > 0;
  }

  const tickets = rows.slice(1)
    .filter(r => r[iReserva] && String(r[iReserva]).startsWith('PB'))
    .map(r => ({
      _pb:        r[iReserva]  || null,
      _date:      parseDate(r[iEmitido]),
      _cancelled: isCancelled(r[iCancel]),
      _operator:  r[iOperador] || null,
      _origin:    r[iOrigen]   || null,
      _dest:      r[iDestino]  || null,
      _price:     parsePrice(r[iPrecio]),
      _priceAPI:  typeof r[iPrecioAPI]==='number' ? r[iPrecioAPI] : parsePrice(r[iPrecioAPI]),
      _gateway:   r[iPGType] && String(r[iPGType]).trim() ? String(r[iPGType]).trim() : null,
      _channel:   r[iMovil] && String(r[iMovil]).trim() !== '-' ? 'Mobile' : 'Desktop',
      _platform:  r[iPlatform] || null,
      _name:      r[iCliente]  || null,
      _email:     r[iCorreo]   || null,
    }));

  const dayMap = {};
  tickets.forEach(t => {
    if (!t._date) return;
    if (!dayMap[t._date]) dayMap[t._date] = { date: t._date, count: 0, cancelled: 0 };
    if (t._cancelled) dayMap[t._date].cancelled++;
    else              dayMap[t._date].count++;
  });

  return {
    byDay:   Object.values(dayMap).sort((a, b) => a.date.localeCompare(b.date)),
    tickets,
  };
}

function filterMainPeriod(rows, dateKey = '_date') {
  const counts = {};
  rows.forEach(r => {
    const d = r[dateKey]; if (d) { const m = d.substring(0,7); counts[m]=(counts[m]||0)+1; }
  });
  const dominant = Object.entries(counts).sort((a,b)=>b[1]-a[1])[0]?.[0];
  return dominant ? rows.filter(r => r[dateKey] && r[dateKey].startsWith(dominant)) : rows;
}

function process(files) {
  const operatorType = files.operatorType || 'konnect';
  const fallidos     = files.fallidos.flatMap(parseCSV);
  const pendientes   = files.pendientes.flatMap(parseCSV);
  const abandonos    = files.abandonos.flatMap(parseCSV);

  let salesByDay     = [];
  let apiTickets     = [];
  let totalCancelled = 0;
  const hasSales     = !!files.sales;

  if (hasSales) {
    if (operatorType === 'api') {
      const parsed    = parseSalesAPI(files.sales);
      salesByDay      = parsed.byDay;
      apiTickets      = parsed.tickets;
      totalCancelled  = salesByDay.reduce((s, d) => s + d.cancelled, 0);
    } else {
      salesByDay = parseSalesKonnect(files.sales);
    }
  }

  const f = filterMainPeriod(fallidos);
  const p = filterMainPeriod(pendientes);
  const a = filterMainPeriod(abandonos);

  // For API: also filter salesByDay to dominant month
  let filteredSales = salesByDay;
  if (hasSales && operatorType === 'api' && salesByDay.length) {
    const mc = {};
    salesByDay.forEach(d => { const m=d.date.substring(0,7); mc[m]=(mc[m]||0)+d.count; });
    const dom = Object.entries(mc).sort((a,b)=>b[1]-a[1])[0]?.[0];
    if (dom) filteredSales = salesByDay.filter(d => d.date.startsWith(dom));
  }

  // Determine dominant period from ALL data to align sales + failures
  const allDates = [
    ...filteredSales.map(s => s.date),
    ...f.map(r => r._date), ...p.map(r => r._date), ...a.map(r => r._date),
  ].filter(Boolean);
  const mc2 = {};
  allDates.forEach(d => { const m=d.substring(0,7); mc2[m]=(mc2[m]||0)+1; });
  const dominantMonth = Object.entries(mc2).sort((a,b)=>b[1]-a[1])[0]?.[0];

  // Build daily map
  const dailyMap = {};
  filteredSales
    .filter(s => !dominantMonth || s.date.startsWith(dominantMonth))
    .forEach(s => { dailyMap[s.date] = { date:s.date, sales:s.count, failures:0, pending:0, abandonments:0 }; });

  [[f,0],[p,1],[a,2]].forEach(([arr,idx]) => arr.forEach(r => {
    if (!r._date) return;
    if (!dailyMap[r._date]) dailyMap[r._date] = { date:r._date, sales:0, failures:0, pending:0, abandonments:0 };
    if      (idx===0) dailyMap[r._date].failures++;
    else if (idx===1) dailyMap[r._date].pending++;
    else              dailyMap[r._date].abandonments++;
  }));

  const daily = Object.values(dailyMap).sort((a,b)=>a.date.localeCompare(b.date)).map(d => {
    const totalAttempts = d.sales + d.failures + d.pending;
    const totalAll      = totalAttempts + d.abandonments;
    const failDenom     = hasSales ? totalAttempts : (d.failures+d.pending+d.abandonments);
    return {
      ...d,
      dateStr:    d.date.split('-').reverse().join('/'),
      failureRate: failDenom>0 ? (d.failures+d.pending)/failDenom*100 : 0,
      abandonRate: totalAll>0  ? d.abandonments/totalAll*100 : 0,
      totalNotConverted: d.failures+d.pending+d.abandonments,
    };
  });

  // Aggregations
  const gwMap = {};
  const addGW = (arr,key) => arr.forEach(r => {
    const g=r._gateway||'Unknown';
    if(!gwMap[g]) gwMap[g]={gateway:g,failures:0,pending:0,abandonments:0};
    gwMap[g][key]++;
  });
  addGW(f,'failures'); addGW(p,'pending'); addGW(a,'abandonments');
  const gateways = Object.values(gwMap).map(g=>({...g,total:g.failures+g.pending+g.abandonments})).sort((a,b)=>b.failures-a.failures||b.total-a.total);
  const totalInc = gateways.reduce((s,g)=>s+g.total,0);
  gateways.forEach(g => g.pct = totalInc ? g.total/totalInc*100 : 0);

  const chMap={};
  f.forEach(r=>{const c=r._channel||'Unknown';if(!chMap[c])chMap[c]={channel:c,failures:0,abandonments:0};chMap[c].failures++;});
  a.forEach(r=>{const c=r._channel||'Unknown';if(!chMap[c])chMap[c]={channel:c,failures:0,abandonments:0};chMap[c].abandonments++;});
  const channels = Object.values(chMap).sort((a,b)=>b.failures-a.failures);

  const plMap={};
  f.forEach(r=>{const pl=r._platform||'Unknown';if(!plMap[pl])plMap[pl]={platform:pl,failures:0,abandonments:0};plMap[pl].failures++;});
  a.forEach(r=>{const pl=r._platform||'Unknown';if(!plMap[pl])plMap[pl]={platform:pl,failures:0,abandonments:0};plMap[pl].abandonments++;});
  const platforms = Object.values(plMap).sort((a,b)=>b.failures-a.failures);

  const totalSales    = hasSales ? daily.reduce((s,d)=>s+d.sales,0) : null;
  const totalFailures = f.length;
  const totalPending  = p.length;
  const totalAbandon  = a.length;
  const totalDays     = daily.length || 1;
  const priceFailures = f.reduce((s,r)=>s+r._price,0);
  const pricePending  = p.reduce((s,r)=>s+r._price,0);
  const priceAbandon  = a.reduce((s,r)=>s+r._price,0);
  const totalLost     = priceFailures+pricePending+priceAbandon;
  const avgSales      = hasSales ? totalSales/totalDays : null;
  const avgFailures   = totalFailures/totalDays;
  const avgAbandon    = totalAbandon/totalDays;
  const avgFailRate   = daily.reduce((s,d)=>s+d.failureRate,0)/totalDays;
  const avgAbanRate   = daily.reduce((s,d)=>s+d.abandonRate,0)/totalDays;
  const peakFailDay   = [...daily].sort((a,b)=>b.failures-a.failures)[0];
  const peakAbanDay   = [...daily].sort((a,b)=>b.abandonments-a.abandonments)[0];

  const dates       = daily.map(d=>d.date);
  const periodStart = dates[0]?.split('-').reverse().join('/') || '—';
  const periodEnd   = dates[dates.length-1]?.split('-').reverse().join('/') || '—';

  return {
    operator:       files.operatorName || 'Bus Operator',
    operatorType,   hasSales,          totalCancelled,
    lang:           files.lang || 'en',
    periodStart,    periodEnd,         totalDays,
    totalSales,     totalFailures,     totalPending,    totalAbandon,
    avgSales,       avgFailures,       avgAbandon,
    avgFailRate,    avgAbanRate,
    priceFailures,  pricePending,      priceAbandon,    totalLost,
    peakFailDay,    peakAbanDay,
    daily, gateways, channels, platforms, totalInc,
    rawFailures: f, rawPending: p, rawAbandon: a, apiTickets,
  };
}

module.exports = { process };