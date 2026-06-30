// ═══════════════════════════════════════════════════════════════
//  POMADEX FINANCE 2.0 — Google Apps Script
//  Sheets ID: 19YutemqC9e3AQz3PR_Mfwn-c7GxgQ1PYd4vAwDRgWmE
//  Estrutura: Data | Tipo | Categoria | Subcategoria | Descrição
//             Valor | Moeda | Conta_Banco | Status | Observacao | Projecto
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '19YutemqC9e3AQz3PR_Mfwn-c7GxgQ1PYd4vAwDRgWmE';

// Abas por projecto
const SHEETS = {
  pp:    'Pomadex_Productions',
  sots:  'SOTS_Recordings',
  snare: 'Clube_do_Snare',
  pemax: 'PEMAX',
  pedro: 'Pedro_Maximo',
};

// Cabeçalhos obrigatórios
const HEADERS = ['Data','Tipo','Categoria','Subcategoria','Descrição','Valor','Moeda','Conta_Banco','Status','Observacao','Projecto'];

// Categorias por projecto
const CATEGORIAS = {
  pp: {
    'Management':      ['Management Fee Mensal','Comissão Booking','Comissão Show','Consultoria'],
    'Produção':        ['Produção Musical','Beatmaking','Arranjo','Composição'],
    'Mixing/Mastering':['Mix Single','Mix EP/Album','Mastering Single','Mastering EP/Album'],
    'DJ Euro Mgmt':    ['Booking Europa','Agência Fee','Rider','Despesas Tour'],
    'Marketing':       ['Meta Ads','Google Ads','Press Release','Fotografia','Vídeo'],
    'Equipamento':     ['Compra Equipamento','Aluguer Equipamento','Manutenção','Software'],
    'Plataformas':     ['Spotify','Apple Music','Beatport','Bandcamp','SoundCloud Pro'],
    'Escritório':      ['Assinaturas','Material','Internet','Telefone'],
  },
  sots: {
    'Distribuição':    ['Royalties Spotify','Royalties Apple','Royalties Beatport','Royalties Bandcamp','Sync License'],
    'Eventos':         ['Booking Show','Cachet DJ Set','Cachet Live','Rider Show'],
    'Produção':        ['Estúdio','Masterização','Artwork','Prensagem Vinil'],
    'Marketing':       ['Campanha Digital','Press Kit','Fotografia','Relações Públicas'],
  },
  snare: {
    'Assinaturas':     ['Membership Mensal','Membership Anual','Membership Trial'],
    'Eventos':         ['Workshop Ticket','Evento Especial','Masterclass'],
    'Conteúdo':        ['Curso Online','Tutorial','Podcast','Merchandise'],
    'Parcerias':       ['Sponsor','Afiliado','Brand Deal'],
  },
  pemax: {
    'DJ Classes':      ['DJ Class 1:1','DJ Class Grupo 2-3','DJ Class Grupo 4+','DJ Class Online','DJ Class Pacote 4','DJ Class Pacote 8'],
    'Shows':           ['Cache Show Local','Cache Show Nacional','Cache Show Europa','Cache Show Online'],
    'Music Classes':   ['Music Production 1:1','Music Production Grupo','Ableton Live Class','Logic Pro Class'],
    'Royalties':       ['Royalties Bandcamp','Royalties Spotify','Sync License','Sample Pack'],
    'Merchandise':     ['T-Shirt','Hoodie','Vinilo','Bundle'],
  },
  pedro: {
    'Habitação':       ['Renda','Electricidade','Gás','Internet','Condomínio'],
    'Alimentação':     ['Supermercado','Restaurante','Take Away','Café'],
    'Transporte':      ['Uber','Táxi','Gasolina','Transporte Público','Parking'],
    'Saúde':           ['Médico','Farmácia','Ginásio','Seguro Saúde'],
    'Pessoal':         ['Roupa','Lazer','Subscrições','Educação','Viagens'],
    'Impostos':        ['IRS','Segurança Social','Contabilista','Outros Impostos'],
  },
};

// ── SETUP: cria abas e cabeçalhos se não existirem ─────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  Object.values(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sh.getRange(1, 1, 1, HEADERS.length)
        .setBackground('#1a1a1a').setFontColor('#f5c842')
        .setFontWeight('bold').setFontSize(11);
      sh.setFrozenRows(1);
      // Column widths
      sh.setColumnWidth(1, 100); // Data
      sh.setColumnWidth(2, 80);  // Tipo
      sh.setColumnWidth(3, 130); // Categoria
      sh.setColumnWidth(4, 150); // Subcategoria
      sh.setColumnWidth(5, 220); // Descrição
      sh.setColumnWidth(6, 90);  // Valor
      sh.setColumnWidth(7, 70);  // Moeda
      sh.setColumnWidth(8, 120); // Conta_Banco
      sh.setColumnWidth(9, 100); // Status
      sh.setColumnWidth(10, 180);// Observacao
      sh.setColumnWidth(11, 100);// Projecto
    }
  });

  // Aba Categorias
  let catSheet = ss.getSheetByName('Categorias');
  if (!catSheet) {
    catSheet = ss.insertSheet('Categorias');
    catSheet.getRange(1,1,1,3).setValues([['Projecto','Categoria','Subcategoria']]);
    catSheet.getRange(1,1,1,3).setBackground('#1a1a1a').setFontColor('#f5c842').setFontWeight('bold');
    let row = 2;
    Object.entries(CATEGORIAS).forEach(([proj, cats]) => {
      Object.entries(cats).forEach(([cat, subs]) => {
        subs.forEach(sub => {
          catSheet.getRange(row, 1, 1, 3).setValues([[proj.toUpperCase(), cat, sub]]);
          row++;
        });
      });
    });
  }

  // Aba Dashboard
  let dash = ss.getSheetByName('Dashboard');
  if (!dash) {
    dash = ss.insertSheet('Dashboard');
    dash.getRange('A1').setValue('Pomadex Finance 2.0 — Dashboard');
    dash.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#f5c842');
    dash.getRange('A3').setValue('Use o app para ver os dados actualizados →');
    dash.getRange('A3').setFontColor('#888888');
  }

  return 'Setup completo!';
}

// ── MAIN HANDLER ──────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    if      (action === 'getTxns')      result = getTxns(body);
    else if (action === 'addTxn')       result = addTxn(body);
    else if (action === 'deleteTxn')    result = deleteTxn(body);
    else if (action === 'importBulk')   result = importBulk(body);
    else if (action === 'getForecasts') result = getForecasts(body);
    else if (action === 'addForecast')  result = addForecast(body);
    else if (action === 'deleteForecast') result = deleteForecast(body);
    else if (action === 'verifyPin')    result = verifyPin(body);
    else if (action === 'changePin')    result = changePin(body);
    else if (action === 'setup')        result = { ok: true, msg: setupSheets() };
    else result = { ok: false, error: 'Unknown action: ' + action };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return doPost(e);
}

// ── GET TRANSACTIONS ──────────────────────────────────────────
function getTxns(body) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const all = [];

  Object.entries(SHEETS).forEach(([projKey, shName]) => {
    const sh = ss.getSheetByName(shName);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || row[2] === '--- FORECAST ---') continue; // skip empty and forecasts
      const tipo = (row[1]||'').toString().toLowerCase();
      if (tipo === 'previsão entrada' || tipo === 'previsão saída') continue;

      const dateVal = row[0];
      let dateStr = '';
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        // Try to parse DD/MM/YYYY
        const parts = dateVal.toString().split('/');
        if (parts.length === 3) dateStr = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
        else dateStr = dateVal.toString();
      }

      all.push({
        id: `${projKey}_${i}`,
        date: dateStr,
        type: tipo === 'entrada' ? 'in' : 'out',
        categoria: row[2] || '',
        subcategoria: row[3] || '',
        desc: row[4] || '',
        amount: parseFloat((row[5]||0).toString().replace(',','.')) || 0,
        moeda: row[6] || 'EUR',
        conta: row[7] || '',
        status: row[8] || '',
        obs: row[9] || '',
        proj: projKey,
        // Legacy fields for app compatibility
        cat: row[3] || row[2] || '',
        description: row[4] || '',
      });
    }
  });

  return { ok: true, data: all };
}

// ── ADD TRANSACTION ───────────────────────────────────────────
function addTxn(body) {
  const txn = body.txn;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shName = SHEETS[txn.proj];
  if (!shName) return { ok: false, error: 'Unknown project: ' + txn.proj };

  let sh = ss.getSheetByName(shName);
  if (!sh) {
    setupSheets();
    sh = ss.getSheetByName(shName);
  }

  // Convert date from yyyy-MM-dd to DD/MM/YYYY
  const dateParts = (txn.date || '').split('-');
  const dateFormatted = dateParts.length === 3
    ? `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`
    : txn.date;

  const tipo = txn.type === 'in' ? 'Entrada' : 'Saída';
  const newId = Date.now();

  sh.appendRow([
    dateFormatted,
    tipo,
    txn.categoria || txn.cat || '',
    txn.subcategoria || txn.cat || '',
    txn.desc || txn.description || '',
    txn.amount || 0,
    txn.moeda || 'EUR',
    txn.conta || '',
    txn.status || 'Confirmado',
    txn.obs || '',
    txn.proj.toUpperCase(),
    newId, // hidden id column
  ]);

  return { ok: true, id: newId };
}

// ── ADD FORECAST ──────────────────────────────────────────────
function addForecast(body) {
  const fc = body.forecast;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shName = SHEETS[fc.proj] || SHEETS.pp;
  let sh = ss.getSheetByName(shName);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(shName); }

  const tipo = fc.type === 'fin' ? 'Previsão Entrada' : 'Previsão Saída';
  const month = (fc.month || 0) + 1;
  const dateStr = `01/${String(month).padStart(2,'0')}/${fc.year || 2026}`;
  const newId = Date.now();

  sh.appendRow([
    dateStr, tipo,
    fc.categoria || fc.cat || '',
    fc.subcategoria || fc.cat || '',
    fc.desc || '',
    fc.amount || 0,
    'EUR', '', 'Previsão', '', fc.proj.toUpperCase(), newId,
  ]);

  return { ok: true, id: newId };
}

// ── GET FORECASTS ─────────────────────────────────────────────
function getForecasts(body) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const all = [];

  Object.entries(SHEETS).forEach(([projKey, shName]) => {
    const sh = ss.getSheetByName(shName);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const tipo = (row[1]||'').toString().toLowerCase();
      if (tipo !== 'previsão entrada' && tipo !== 'previsão saída') continue;

      const dateVal = row[0].toString();
      const parts = dateVal.split('/');
      const month = parts.length >= 2 ? parseInt(parts[1]) - 1 : 0;
      const year  = parts.length >= 3 ? parseInt(parts[2]) : 2026;

      all.push({
        id: `fc_${projKey}_${i}`,
        type: tipo === 'previsão entrada' ? 'fin' : 'fout',
        month, year,
        proj: projKey,
        cat: row[3] || row[2] || '',
        desc: row[4] || '',
        amount: parseFloat((row[5]||0).toString().replace(',','.')) || 0,
        status: row[8] || 'Previsão',
      });
    }
  });

  return { ok: true, data: all };
}

// ── DELETE ────────────────────────────────────────────────────
function deleteTxn(body) {
  // Find and delete row by id (last column)
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const idStr = body.id.toString();

  for (const shName of Object.values(SHEETS)) {
    const sh = ss.getSheetByName(shName);
    if (!sh) continue;
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][11] && data[i][11].toString() === idStr) {
        sh.deleteRow(i + 1);
        return { ok: true };
      }
    }
  }
  return { ok: false, error: 'Not found' };
}

function deleteForecast(body) {
  return deleteTxn(body); // same logic
}

// ── IMPORT BULK ───────────────────────────────────────────────
function importBulk(body) {
  const txns = body.txns || [];
  let count = 0;

  txns.forEach(txn => {
    try {
      addTxn({ txn });
      count++;
    } catch(e) {}
  });

  return { ok: true, count };
}

// ── PIN ───────────────────────────────────────────────────────
function verifyPin(body) {
  const store = PropertiesService.getScriptProperties();
  const saved = store.getProperty('PIN') || '1234';
  return { ok: true, verified: body.pin === saved };
}

function changePin(body) {
  const store = PropertiesService.getScriptProperties();
  const saved = store.getProperty('PIN') || '1234';
  if (body.oldPin !== saved) return { ok: false, error: 'Wrong PIN' };
  store.setProperty('PIN', body.newPin);
  return { ok: true };
}

