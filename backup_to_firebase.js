const fs = require('fs');
const https = require('https');
const path = require('path');

function getFirebaseToken() {
  const configPath = path.join(process.env.USERPROFILE, '.config', 'configstore', 'firebase-tools.json');
  if (!fs.existsSync(configPath)) {
    throw new Error('No firebase-tools.json found at ' + configPath);
  }
  const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  return config.tokens;
}

function refreshAccessToken(tokens) {
  return new Promise((resolve, reject) => {
    // If access token is valid, we can try using it or refresh via googleapis
    const postData = new URLSearchParams({
      client_id: '563584335869-fgrhgmd47bqnekij5i8b5pr03ho85od6.apps.googleusercontent.com',
      grant_type: 'refresh_token',
      refresh_token: tokens.refresh_token
    }).toString();

    const req = https.request({
      hostname: 'oauth2.googleapis.com',
      path: '/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.access_token) {
            resolve(json.access_token);
          } else {
            // fallback to tokens.access_token
            resolve(tokens.access_token);
          }
        } catch (e) {
          resolve(tokens.access_token);
        }
      });
    });
    req.on('error', () => resolve(tokens.access_token));
    req.write(postData);
    req.end();
  });
}

function firestoreSetDocument(token, collection, docId, fields) {
  return new Promise((resolve, reject) => {
    const postData = JSON.stringify({ fields });
    const reqPath = `/v1/projects/mantenimiento-21758/databases/presupuesto-2026-backup/documents/${collection}/${docId}`;

    const req = https.request({
      hostname: 'firestore.googleapis.com',
      path: reqPath,
      method: 'PATCH',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(postData)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(JSON.parse(data));
        } else {
          reject(new Error(`HTTP ${res.statusCode}: ${data}`));
        }
      });
    });
    req.on('error', reject);
    req.write(postData);
    req.end();
  });
}

async function runBackup() {
  console.log('--- Iniciando Respaldo en Firebase Firestore: presupuesto-2026-backup ---');
  const tokens = getFirebaseToken();
  const accessToken = await refreshAccessToken(tokens);

  const dataFile = path.join(__dirname, 'web_dashboard', 'dashboard_data.json');
  if (!fs.existsSync(dataFile)) {
    throw new Error('dashboard_data.json no encontrado');
  }

  const raw = fs.readFileSync(dataFile, 'utf8');
  const dataset = JSON.parse(raw);

  const timestamp = new Date().toISOString();

  // 1. Guardar metadatos generales del backup
  console.log('1. Guardando metadatos generales...');
  await firestoreSetDocument(accessToken, 'metadata', 'latest', {
    backup_timestamp: { stringValue: timestamp },
    database_target: { stringValue: 'presupuesto-2026-backup' },
    project: { stringValue: 'mantenimiento-21758' },
    source: { stringValue: 'dashboard_data.json' },
    periods: {
      arrayValue: {
        values: Object.keys(dataset.periods || {}).map(p => ({ stringValue: p }))
      }
    },
    version: { stringValue: '2.5.0-multilang' }
  });

  // 2. Guardar resumen por cada período
  for (const [periodKey, periodData] of Object.entries(dataset.periods || {})) {
    console.log(`2. Guardando resumen para período ${periodKey}...`);
    const kpis = periodData.global_kpis || {};
    const comObj = periodData.comerciales || {};
    const comNames = Object.keys(comObj);

    await firestoreSetDocument(accessToken, 'periods', periodKey, {
      period_id: { stringValue: periodKey },
      nombre: { stringValue: periodData.nombre || periodKey },
      fecha_corte: { stringValue: kpis.fecha_corte || '' },
      budget_uds: { doubleValue: Number(kpis.budget_uds || 0) },
      pedidos_uds: { doubleValue: Number(kpis.pedidos_uds || 0) },
      servidas_uds: { doubleValue: Number(kpis.servidas_uds || 0) },
      pendientes_uds: { doubleValue: Number(kpis.pendientes_uds || 0) },
      gap_uds: { doubleValue: Number(kpis.gap_uds || 0) },
      pct_consec: { doubleValue: Number(kpis.pct_consec || 0) },
      forecast_accuracy: { doubleValue: Number(kpis.forecast_accuracy || 0) },
      budget_nacional: { doubleValue: Number(kpis.budget_uds_nacional || 0) },
      pedidos_nacional: { doubleValue: Number(kpis.pedidos_uds_nacional || 0) },
      comerciales_count: { integerValue: comNames.length.toString() },
      backup_timestamp: { stringValue: timestamp }
    });

    // Guardar resumen comercial
    for (const comName of comNames) {
      const cData = comObj[comName];
      const cKpis = cData.kpis || {};
      const docName = `${periodKey}_${comName.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]/g, '_')}`;
      await firestoreSetDocument(accessToken, 'comerciales', docName, {
        period_id: { stringValue: periodKey },
        comercial: { stringValue: comName },
        budget_uds: { doubleValue: Number(cKpis.budget_uds || 0) },
        pedidos_uds: { doubleValue: Number(cKpis.pedidos_uds || 0) },
        servidas_uds: { doubleValue: Number(cKpis.servidas_uds || 0) },
        pendientes_uds: { doubleValue: Number(cKpis.pendientes_uds || 0) },
        gap_uds: { doubleValue: Number(cKpis.gap_uds || 0) },
        pct_consec: { doubleValue: Number(cKpis.pct_consec || 0) },
        lineas_count: { integerValue: (cData.lineas || []).length.toString() }
      });
    }
  }

  // 3. Guardar log histórico de respaldo
  const logId = `backup_${timestamp.replace(/[^0-9]/g, '').substring(0, 14)}`;
  await firestoreSetDocument(accessToken, 'backup_history', logId, {
    timestamp: { stringValue: timestamp },
    status: { stringValue: 'SUCCESS' },
    size_bytes: { integerValue: Buffer.byteLength(raw).toString() }
  });

  console.log('✅ Respaldo completado con éxito en Firebase: presupuesto-2026-backup');
}

runBackup().catch(err => {
  console.error('❌ Error en el respaldo:', err);
  process.exit(1);
});
