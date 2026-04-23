'use strict';
const express  = require('express');
const multer   = require('multer');
const path     = require('path');
const { process: processData } = require('./src/processor');
const excelGen = require('./src/excelGen');
const wordGen  = require('./src/wordGen');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });
app.use(express.static(path.join(__dirname, 'public')));

const fields = [
  { name: 'fallidos',   maxCount: 20 },
  { name: 'pendientes', maxCount: 20 },
  { name: 'abandonos',  maxCount: 20 },
  { name: 'sales',      maxCount: 1  },
];

app.post('/generate', upload.fields(fields), async (req, res) => {
  try {
    const operatorName = (req.body.operator || 'Bus Operator').trim();
    const operatorType = (req.body.operatorType || 'konnect').trim();
    const lang = (req.body.lang || 'en').trim();
    const files = req.files || {};

    const salesBuffer = files.sales?.[0]?.buffer || null;
    const hasCsvs = (files.fallidos?.length || 0) + (files.pendientes?.length || 0) + (files.abandonos?.length || 0) > 0;
    if (!hasCsvs && !salesBuffer) {
      return res.status(400).json({ error: 'Upload at least one CSV file or the sales Excel.' });
    }

    const data = processData({
      fallidos: (files.fallidos || []).map(f => f.buffer),
      pendientes: (files.pendientes || []).map(f => f.buffer),
      abandonos: (files.abandonos || []).map(f => f.buffer),
      sales: salesBuffer,
      operatorName,
      operatorType,
      lang,
    });

    // Validate required keys
    const requiredKeys = ['periodStart', 'periodEnd'];
    for (const key of requiredKeys) {
      if (!(key in data)) {
        return res.status(400).json({ error: `Missing required key: ${key}` });
      }
    }

    console.log('Data processed successfully:', data);

    const [xlBuf, docBuf] = await Promise.all([
      Promise.resolve(excelGen.generate(data)),
      wordGen.generate(data),
    ]);

    const slug = operatorName.replace(/[^a-z0-9]/gi, '_');
    const date = new Date().toISOString().slice(0, 10);

    res.json({
      ok: true,
      stats: {
        operator: data.operator,
        operatorType: data.operatorType,
        hasSales: data.hasSales,
        totalCancelled: data.totalCancelled,
        period: `${data.periodStart} – ${data.periodEnd}`,
        days: data.totalDays,
        sales: data.totalSales,
        failures: data.totalFailures,
        pending: data.totalPending,
        abandonments: data.totalAbandon,
        lostRevenue: data.totalLost,
        avgFailRate: data.avgFailRate,
        avgAbanRate: data.avgAbanRate,
        topGateway: data.gateways[0]?.gateway || '—',
        topGwFailures: data.gateways[0]?.failures || 0,
        daily: data.daily,
        gateways: data.gateways,
        channels: data.channels,
        platforms: data.platforms,
      },
      files: {
        excel: { data: xlBuf.toString('base64'), name: `Report_${slug}_${date}.xlsx` },
        word: { data: docBuf.toString('base64'), name: `Report_${slug}_${date}.docx` },
      },
    });
  } catch (err) {
    console.error('Error in /generate endpoint:', err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`\n🚌  Bus Reports Portal → http://localhost:${PORT}\n`));