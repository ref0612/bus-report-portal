'use strict';
const { spawnSync } = require('child_process');
const path = require('path');

const PY_SCRIPT = path.join(__dirname, 'excelGen.py');

function generate(data) {
  const payload = {
    ...data,
    rawFailures: (data.rawFailures || []).map(r => ({
      _pb:r._pb,_date:r._date,_origin:r._origin,_dest:r._dest,_seat:r._seat,
      _operator:r._operator,_channel:r._channel,_gateway:r._gateway,
      _pgStatus:r._pgStatus,_platform:r._platform,_price:r._price,
    })),
    rawPending: (data.rawPending || []).map(r => ({
      _pb:r._pb,_date:r._date,_origin:r._origin,_dest:r._dest,_seat:r._seat,
      _operator:r._operator,_channel:r._channel,_gateway:r._gateway,
      _price:r._price,_email:r._email,
    })),
    rawAbandon: (data.rawAbandon || []).map(r => ({
      _pb:r._pb,_date:r._date,_origin:r._origin,_dest:r._dest,_seat:r._seat,
      _operator:r._operator,_channel:r._channel,_gateway:r._gateway,
      _pgStatus:r._pgStatus,_platform:r._platform,_price:r._price,
    })),
    today: new Date().toLocaleDateString('en-GB'),
    apiTickets: [],
  };

  const inputBuf = Buffer.from(JSON.stringify(payload), 'utf-8');
  const result = spawnSync('python3', [PY_SCRIPT], {
    input: inputBuf,
    maxBuffer: 150 * 1024 * 1024,
  });

  if (result.error) throw result.error;
  if (result.status !== 0) {
    const stderr = result.stderr ? result.stderr.toString() : 'unknown error';
    throw new Error(`Excel generator failed (exit ${result.status}): ${stderr.slice(0, 500)}`);
  }

  // stdout is base64-encoded xlsx bytes
  const b64 = result.stdout.toString('utf-8').trim();
  return Buffer.from(b64, 'base64');
}

module.exports = { generate };
