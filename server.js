process.env.TZ = 'America/Sao_Paulo';
try { require('dotenv').config(); } catch(_) {}
const express  = require('express');
const fs       = require('fs');
const path     = require('path');
const https    = require('https');
const crypto   = require('crypto');
const session  = require('express-session');
const pdfParse = require('pdf-parse');
const XLSX     = require('xlsx');

const app          = express();
const DATA_DIR     = path.join(__dirname, 'data');
const TOKENS_FILE  = path.join(DATA_DIR, 'tokens.json');
const RESULT_FILE  = path.join(DATA_DIR, 'gerencial.json');
const USERS_FILE   = path.join(DATA_DIR, 'users.json');
const SHARED_DRIVE    = process.env.SHARED_DRIVE_ID    || '0AKZcsytstd78Uk9PVA';
const EVENTOS_FOLDER  = process.env.EVENTOS_FOLDER_ID  || '1OjS3q7vAccft_n4novmv6d86MBrwiQ9k';
const CLIENT_ID    = process.env.GOOGLE_CLIENT_ID    || '';
const CLIENT_SECRET= process.env.GOOGLE_CLIENT_SECRET|| '';
const PORT         = process.env.PORT || 3001;
const BASE_URL     = process.env.BASE_URL || `http://localhost:${PORT}`;
const REDIRECT_URI = `${BASE_URL}/auth/callback`;
const SESSION_SECRET = process.env.SESSION_SECRET || 'bacco-gerencial-secret-2026';

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// ── Usuários ──────────────────────────────────────────────────────────────────
function hashPwd(pwd) { return crypto.createHash('sha256').update(pwd + 'bacco-salt').digest('hex'); }

function loadUsers() {
  if (!fs.existsSync(USERS_FILE)) {
    const defaults = {
      andrea: { displayName: 'Andrea', password: hashPwd('1234'), mustChange: true },
      yoshio: { displayName: 'Yoshio', password: hashPwd('1234'), mustChange: true }
    };
    fs.writeFileSync(USERS_FILE, JSON.stringify(defaults, null, 2));
    return defaults;
  }
  return JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
}

function saveUsers(users) { fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2)); }

// ── Sessão e middleware ───────────────────────────────────────────────────────
app.use(session({ secret: SESSION_SECRET, resave: false, saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 } }));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));

const PUBLIC_PATHS = ['/login', '/login.html', '/trocar-senha', '/trocar-senha.html'];
app.use((req, res, next) => {
  const isPublic = PUBLIC_PATHS.some(p => req.path.startsWith(p)) || req.path.startsWith('/auth/');
  if (!req.session?.user && !isPublic) return res.redirect('/login');
  next();
});

app.use(express.static(path.join(__dirname, 'public')));

// ── Rotas de autenticação de usuário ─────────────────────────────────────────
app.get('/login', (req, res) => res.sendFile(path.join(__dirname, 'public', 'login.html')));
app.get('/trocar-senha', (req, res) => res.sendFile(path.join(__dirname, 'public', 'trocar-senha.html')));

app.post('/login', (req, res) => {
  const { usuario, senha } = req.body;
  const users = loadUsers();
  const user = users[usuario?.toLowerCase()];
  if (!user || user.password !== hashPwd(senha)) {
    return res.redirect('/login?erro=1');
  }
  req.session.user = { id: usuario.toLowerCase(), name: user.displayName, mustChange: user.mustChange };
  if (user.mustChange) return res.redirect('/trocar-senha');
  res.redirect('/');
});

app.post('/trocar-senha', (req, res) => {
  if (!req.session?.user) return res.redirect('/login');
  const { nova, confirma } = req.body;
  if (!nova || nova.length < 6) return res.redirect('/trocar-senha?erro=curta');
  if (nova !== confirma)        return res.redirect('/trocar-senha?erro=naoconfere');
  const users = loadUsers();
  users[req.session.user.id].password   = hashPwd(nova);
  users[req.session.user.id].mustChange = false;
  saveUsers(users);
  req.session.user.mustChange = false;
  res.redirect('/');
});

app.get('/logout', (req, res) => { req.session.destroy(); res.redirect('/login'); });

// ── HTTP helper ───────────────────────────────────────────────────────────────
function req(url, opts = {}, body = null) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const options = {
      hostname: u.hostname, path: u.pathname + u.search,
      method: opts.method || 'GET',
      headers: { 'User-Agent': 'bacco-gerencial/1.0', ...opts.headers }
    };
    const r = https.request(options, res => {
      if (res.statusCode === 301 || res.statusCode === 302)
        return req(res.headers.location, opts).then(resolve).catch(reject);
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks) }));
      res.on('error', reject);
    });
    r.on('error', reject);
    r.setTimeout(60000, () => { r.destroy(); reject(new Error('timeout')); });
    if (body) r.write(body);
    r.end();
  });
}

// ── OAuth2 ────────────────────────────────────────────────────────────────────
const loadTokens = () => { try { return JSON.parse(fs.readFileSync(TOKENS_FILE,'utf8')); } catch { return null; } };
const saveTokens = t => fs.writeFileSync(TOKENS_FILE, JSON.stringify(t, null, 2));

async function getToken() {
  let t = loadTokens();
  if (!t?.refresh_token) throw new Error('NÃO_AUTORIZADO');
  if (!t.expiry || Date.now() > t.expiry - 60000) {
    const body = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET,
      refresh_token: t.refresh_token, grant_type: 'refresh_token' }).toString();
    const { status, body: rb } = await req('https://oauth2.googleapis.com/token',
      { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }, body);
    const d = JSON.parse(rb.toString());
    if (status !== 200) throw new Error('Token inválido: ' + (d.error_description || d.error));
    t = { ...t, access_token: d.access_token, expiry: Date.now() + d.expires_in * 1000 };
    saveTokens(t);
  }
  return t.access_token;
}

// ── Drive API ─────────────────────────────────────────────────────────────────
async function driveGet(endpoint) {
  const token = await getToken();
  const sep = endpoint.includes('?') ? '&' : '?';
  const url = `https://www.googleapis.com/drive/v3/${endpoint}${sep}supportsAllDrives=true&includeItemsFromAllDrives=true`;
  const { status, body } = await req(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = JSON.parse(body.toString());
  if (status >= 300) throw new Error(data.error?.message || `Drive API HTTP ${status}`);
  return data;
}

async function findFolder(parentId, name) {
  const q = encodeURIComponent(`'${parentId}' in parents and name='${name}' and mimeType='application/vnd.google-apps.folder' and trashed=false`);
  const d = await driveGet(`files?q=${q}&fields=files(id,name)&corpora=allDrives`);
  return d.files?.[0] || null;
}

async function allPdfs(folderId) {
  const q = encodeURIComponent(`'${folderId}' in parents and mimeType='application/pdf' and trashed=false`);
  const d = await driveGet(`files?q=${q}&fields=files(id,name,modifiedTime)&orderBy=name&corpora=allDrives`);
  return d.files || [];
}

async function latestPdf(folderId) {
  const files = await allPdfs(folderId);
  return files.sort((a,b) => b.modifiedTime.localeCompare(a.modifiedTime))[0] || null;
}

// Nomes completos primeiro (para evitar que "MAI" case em "MAIO" antes de "MAR" em "MARÇO")
const MESES_PT = [
  'JANEIRO','FEVEREIRO','MARÇO','MARCO','ABRIL','MAIO','JUNHO',
  'JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO',
  // Abreviações (verificadas depois dos nomes completos)
  'JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'
];
const MESES_NUM = {
  JANEIRO:1, JAN:1,
  FEVEREIRO:2, FEV:2,
  'MARÇO':3, MARCO:3, MAR:3,
  ABRIL:4, ABR:4,
  MAIO:5, MAI:5,
  JUNHO:6, JUN:6,
  JULHO:7, JUL:7,
  AGOSTO:8, AGO:8,
  SETEMBRO:9, SET:9,
  OUTUBRO:10, OUT:10,
  NOVEMBRO:11, NOV:11,
  DEZEMBRO:12, DEZ:12
};

function mesKey(filename) {
  // Normaliza: remove acentos, extensão, espaços
  const up = filename.toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\.(PDF|XLSX?)$/,'');
  for (const m of MESES_PT) {
    if (up.includes(m)) {
      const yr = (up.match(/\d{4}/) || [''])[0];
      const num = String(MESES_NUM[m]).padStart(2,'0');
      return yr ? `${yr}-${num}` : null;
    }
  }
  return null;
}

function mesLabel(key) {
  const [yr, mo] = key.split('-');
  const nomes = ['','Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                 'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return `${nomes[parseInt(mo)]} ${yr}`;
}

async function downloadFile(fileId) {
  const token = await getToken();
  const { status, body } = await req(
    `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media&supportsAllDrives=true`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (status !== 200) throw new Error(`Download falhou HTTP ${status}`);
  return body;
}

// ── Parsers ───────────────────────────────────────────────────────────────────
function parseVendas(text) {
  const lines = text.split('\n');
  let pdv = 'RESTAURANTE';
  const notas = { RESTAURANTE: new Set(), 'Room Service': new Set() };
  const daily = {};
  const brutoTaxa = { RESTAURANTE: 0, 'Room Service': 0 };
  let lastDate = null;
  let grand = null;

  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];
    const line = raw.trim();
    if (/^PDV:\s*(RESTAURANTE|Room Service)$/.test(line)) {
      pdv = line.replace(/^PDV:\s*/, '').trim(); continue;
    }
    const dm = line.match(/^(\d{2}\/\d{2}\/\d{4})/);
    if (dm) lastDate = dm[1];
    const nm = line.match(/(\d{5,6})\s*$/);
    if (nm && lastDate && !line.startsWith('TOTAL')) notas[pdv]?.add(nm[1]);

    if (line.startsWith('TOTAL DO DIA:') && lastDate) {
      // Estrutura c/ QTD decimal: QTD | BRUTO | DESCONTO | LIQUIDO | TAXA | TOTAL | CUSTO (7 vals)
      // Estrutura c/ QTD inteiro (RS): BRUTO | DESCONTO | LIQUIDO | TAXA | TOTAL | CUSTO (6 vals)
      // bruto + taxa = totalPago + desconto (identidade matemática, independe do formato)
      const nums = [...line.matchAll(/[\d]+[.,][\d]+/g)].map(m => parseFloat(m[0].replace(/\./g,'').replace(',','.')));
      if (nums.length >= 5) {
        const tp       = nums[nums.length - 2];
        const desconto = nums.length >= 7 ? nums[2]
                       : nums.length >= 6 ? nums[1]
                       : 0;
        if (!daily[lastDate]) daily[lastDate] = { RESTAURANTE: 0, 'Room Service': 0 };
        daily[lastDate][pdv] = (daily[lastDate][pdv] || 0) + tp;
        brutoTaxa[pdv] = (brutoTaxa[pdv] || 0) + tp + desconto;
      }
    } else if (line.startsWith('TOTAL:') && !line.startsWith('TOTAL DO')) {
      // Totalizador geral: linha pode estar quebrada em múltiplas linhas no PDF
      // Junta as próximas 5 linhas para capturar todos os valores
      const bloco = lines.slice(i, i + 6).join(' ');
      const nums = [...bloco.matchAll(/[\d]+[.,][\d]+/g)].map(m => parseFloat(m[0].replace(/\./g,'').replace(',','.')));
      if (nums.length >= 7) {
        grand = {
          valorBruto:   nums[2],
          desconto:     nums[3],
          valorLiquido: nums[4],
          taxaServico:  nums[5],
          totalPago:    nums[6]
        };
      }
    }
  }
  return { daily, notas: { RESTAURANTE: notas.RESTAURANTE.size, 'Room Service': notas['Room Service'].size }, brutoTaxa, grand };
}

function parseOcupacao(text) {
  const lines = text.split('\n');
  const daily = {};
  let currentDate = null;
  let prevLine = '';
  let total = 0;
  let receitaTotal = 0;
  for (const raw of lines) {
    const line = raw.trim();
    // "Data Lançamento" is alone, next non-empty line is the date
    if (/^Data\s+Lan/i.test(line)) { currentDate = null; prevLine = 'HEADER'; continue; }
    if (prevLine === 'HEADER' && /^\d{2}\/\d{2}\/\d{4}$/.test(line)) {
      currentDate = line; prevLine = ''; continue;
    }
    // Summary line before "Total por Data:" — format: "1.320,006600" → valor(R$) + ADs + 00
    if (line === 'Total por Data:' && currentDate && prevLine) {
      const m = prevLine.match(/^([\d.]+,\d{2})(\d+?)00$/);
      if (m) {
        const ad  = parseInt(m[2]);
        const brl = parseFloat(m[1].replace(/\./g,'').replace(',','.'));
        daily[currentDate] = { hospedes: ad, receita: brl };
        total      += ad;
        receitaTotal += brl;
      }
      currentDate = null;
    }
    if (line) prevLine = line;
  }
  return { daily, total, receitaTotal };
}

// ── Parser de Eventos (planilha xlsx) ────────────────────────────────────────
const SHEET_MES = {
  'JANEIRO':'2026-01','FEVEREIRO':'2026-02','MARÇO':'2026-03','MARCO':'2026-03',
  'ABRIL':'2026-04','MAIO':'2026-05','JUNHO':'2026-06','JULHO':'2026-07',
  'AGOSTO':'2026-08','SETEMBRO':'2026-09','OUTUBRO':'2026-10',
  'NOVEMBRO':'2026-11','DEZEMBRO':'2026-12'
};

function parseEventos(buffer) {
  const wb   = XLSX.read(buffer, { type: 'buffer' });
  const result = {};
  for (const sheetName of wb.SheetNames) {
    const mesKey = SHEET_MES[sheetName.toUpperCase().trim().normalize('NFD').replace(/[̀-ͯ]/g,'')];
    if (!mesKey) continue;
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });

    // Encontra a linha de cabeçalho procurando por "PAX"
    let hRow = -1, cPax = -1, cBanq = -1, cForma = -1, cSala = -1, cEquip = -1, cData = -1;
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const r = rows[i].map(c => String(c).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,'').trim());
      const pi = r.findIndex(c => c.includes('PAX'));
      if (pi >= 0) {
        hRow  = i; cPax  = pi;
        cBanq = r.findIndex(c => c.startsWith('BAN'));
        cSala = r.findIndex(c => c.startsWith('SAL'));
        cEquip= r.findIndex(c => c.startsWith('EQUI'));
        cForma= r.findIndex(c => c.includes('FORMA') || (c.includes('PAGAMENTO') && c.length > 10));
        cData = r.findIndex(c => c === 'DATA' || c === 'DT' || c.startsWith('DATA'));
        break;
      }
    }
    if (hRow < 0 || cBanq < 0) continue;

    const parseVal = v => typeof v === 'number' ? v
      : (parseFloat(String(v).replace(/[R$\s]/g,'').replace(/\./g,'').replace(',','.')) || 0);

    // Converte valor de data Excel (serial ou string) para dd/mm/yyyy
    const parseDateCell = v => {
      if (!v) return null;
      if (typeof v === 'number') {
        // Serial Excel → JS Date
        const d = new Date(Math.round((v - 25569) * 86400 * 1000));
        const dd = String(d.getUTCDate()).padStart(2,'0');
        const mm = String(d.getUTCMonth()+1).padStart(2,'0');
        const yyyy = d.getUTCFullYear();
        return `${dd}/${mm}/${yyyy}`;
      }
      const s = String(v).trim();
      // Já no formato dd/mm/yyyy ou dd/mm/yy
      if (/^\d{2}\/\d{2}\/\d{2,4}$/.test(s)) {
        const parts = s.split('/');
        if (parts[2].length === 2) parts[2] = '20' + parts[2];
        return parts.join('/');
      }
      return null;
    };

    let totalPax = 0, totalBanq = 0, totalSala = 0, totalEquip = 0;
    const daily = {};
    for (let i = hRow + 1; i < rows.length; i++) {
      const row     = rows[i];
      const evento  = String(row[0] || '').trim();
      const paxRaw  = row[cPax];
      const banqRaw = row[cBanq];
      const forma   = String(row[cForma] || '').toUpperCase();

      // Para ao encontrar linha completamente em branco (sem evento e sem PAX)
      if (!evento && (paxRaw === '' || paxRaw === null || paxRaw === undefined)) break;

      if (paxRaw === '' || paxRaw === null || paxRaw === undefined) continue;
      const pax = parseInt(paxRaw);
      if (!pax || isNaN(pax) || pax <= 0) continue;
      if (forma.includes('BACCO')) continue;

      const banq  = parseVal(banqRaw);
      const sala  = cSala  >= 0 ? parseVal(row[cSala])  : 0;
      const equip = cEquip >= 0 ? parseVal(row[cEquip]) : 0;
      const rowTotal = +(sala + equip + banq).toFixed(2);

      totalPax  += pax;
      totalBanq += banq;
      totalSala += sala;
      totalEquip+= equip;

      // Acumula por data se coluna DATA disponível
      const dateKey = cData >= 0 ? parseDateCell(row[cData]) : null;
      if (dateKey) {
        if (!daily[dateKey]) daily[dateKey] = { pax:0, sala:0, equip:0, banq:0, total:0 };
        daily[dateKey].pax   += pax;
        daily[dateKey].sala  += sala;
        daily[dateKey].equip += equip;
        daily[dateKey].banq  += banq;
        daily[dateKey].total += rowTotal;
      }
    }
    const total = +(totalSala + totalEquip + totalBanq).toFixed(2);
    result[mesKey] = { pax: totalPax, sala: +totalSala.toFixed(2), equip: +totalEquip.toFixed(2), banq: +totalBanq.toFixed(2), total, daily };
  }
  return result;
}

async function findEventosXlsx() {
  // Busca na pasta compartilhada e também em My Drive
  const q = encodeURIComponent(`'${EVENTOS_FOLDER}' in parents and trashed=false and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel')`);
  const d = await driveGet(`files?q=${q}&fields=files(id,name,modifiedTime)&corpora=allDrives`);
  return d.files?.[0] || null;
}

// ── Sincronização ─────────────────────────────────────────────────────────────
function buildMesData(vendaPdf, ocupPdf, vendas, ocupacao, eventosmes) {
  const totalRSTdiario = Object.values(vendas.daily).reduce((s,d) => s + (d.RESTAURANTE||0), 0);
  const totalRSdiario  = Object.values(vendas.daily).reduce((s,d) => s + (d['Room Service']||0), 0);
  // Usa bruto+taxa por PDV (acumulado dos TOTAL DO DIA), com fallback nos totais diários
  const totalRST  = vendas.brutoTaxa?.RESTAURANTE  || totalRSTdiario;
  const totalRS   = vendas.brutoTaxa?.['Room Service'] || totalRSdiario;
  const totalPago = +(totalRST + totalRS).toFixed(2);
  const clientes  = vendas.notas.RESTAURANTE + vendas.notas['Room Service'];
  const hospedes  = ocupacao.total;

  const allDates = [...new Set([...Object.keys(vendas.daily), ...Object.keys(ocupacao.daily)])].sort();
  // Usa breakdown diário de eventos se disponível; caso contrário distribui igualmente
  const eventosDailyMap = eventosmes?.daily || {};
  const temDailyEventos = Object.keys(eventosDailyMap).length > 0;
  const eventosDiarioFallback = (!temDailyEventos && allDates.length > 0)
    ? +(receitaEventos / allDates.length).toFixed(2) : 0;

  const serie = allDates.map(d => ({
    data:        d,
    restaurante: vendas.daily[d]?.RESTAURANTE || 0,
    roomService: vendas.daily[d]?.['Room Service'] || 0,
    cafe:        ocupacao.daily[d]?.receita  || 0,
    eventos:     temDailyEventos ? (eventosDailyMap[d]?.total || 0) : eventosDiarioFallback,
    totalDia:   (vendas.daily[d]?.RESTAURANTE || 0) + (vendas.daily[d]?.['Room Service'] || 0),
    hospedes:    ocupacao.daily[d]?.hospedes || 0,
    receitaCafe: ocupacao.daily[d]?.receita  || 0
  }));

  const clientesCafe    = hospedes;
  const receitaCafe     = +ocupacao.receitaTotal.toFixed(2);
  const clientesEventos = eventosmes?.pax   || 0;
  const eventosSala     = eventosmes?.sala  || 0;
  const eventosEquip    = eventosmes?.equip || 0;
  const eventosBanq     = eventosmes?.banq  || 0;
  const receitaEventos  = eventosmes?.total || eventosBanq;
  const totalGeral      = +(totalPago + receitaCafe + receitaEventos).toFixed(2);
  const totalClientes   = clientes + clientesCafe + clientesEventos;
  const ticketCafe      = clientesCafe    > 0 ? Math.round(receitaCafe    / clientesCafe)    : 0;
  const ticketEventos   = clientesEventos > 0 ? Math.round(receitaEventos / clientesEventos) : 0;
  const ticketGeral     = totalClientes   > 0 ? Math.round(totalGeral     / totalClientes)   : 0;
  const kpiCobertura    = hospedes > 0 ? vendas.notas['Room Service'] / hospedes : 0;

  return {
    arquivoVendas:        vendaPdf?.name || '',
    arquivoOcupacao:      ocupPdf?.name  || '',
    periodo:              allDates.length ? `${allDates[0]} a ${allDates[allDates.length-1]}` : '',
    diasComDados:         allDates.length,
    // Faturamento por canal
    faturamentoRST:       +totalRST.toFixed(2),
    faturamentoRS:        +totalRS.toFixed(2),
    receitaCafe,
    receitaEventos,
    eventosSala,
    eventosEquip,
    eventosBanq,
    faturamentoTotal:     totalGeral,
    // Clientes por canal
    clientesBacco:        vendas.notas.RESTAURANTE,
    clientesRoomService:  vendas.notas['Room Service'],
    clientesCafe,
    clientesEventos,
    clientesTotal:        totalClientes,
    // Tickets por canal
    ticketRST:            vendas.notas.RESTAURANTE > 0 ? Math.round(totalRST/vendas.notas.RESTAURANTE) : 0,
    ticketRS:             vendas.notas['Room Service'] > 0 ? Math.round(totalRS/vendas.notas['Room Service']) : 0,
    ticketCafe,
    ticketEventos,
    ticketGeral,
    // Legado / resumo financeiro
    hospedes,
    ticketMedio:          ticketGeral,
    valorBruto:           vendas.grand?.valorBruto  || 0,
    desconto:             vendas.grand?.desconto     || 0,
    valorLiquido:         vendas.grand?.valorLiquido || 0,
    taxaServico:          vendas.grand?.taxaServico  || 0,
    _debugGrand:          vendas.grand || null,
    kpiCobertura:         +kpiCobertura.toFixed(3),
    serie
  };
}

async function sincronizar() {
  const [vendaDir, ocupDir] = await Promise.all([
    findFolder(SHARED_DRIVE, 'VENDAS'),
    findFolder(SHARED_DRIVE, 'OCUPAÇÃO')
  ]);
  if (!vendaDir) throw new Error('Pasta VENDAS não encontrada no Drive compartilhado.');
  if (!ocupDir)  throw new Error('Pasta OCUPAÇÃO não encontrada no Drive compartilhado.');

  const [vendaPdfs, ocupPdfs] = await Promise.all([
    allPdfs(vendaDir.id),
    allPdfs(ocupDir.id)
  ]);

  // Group by mes key derived from filename
  const vendaMap = {};
  for (const f of vendaPdfs) { const k = mesKey(f.name); if (k) vendaMap[k] = f; }
  const ocupMap  = {};
  for (const f of ocupPdfs)  { const k = mesKey(f.name); if (k) ocupMap[k]  = f; }

  const meses = [...new Set([...Object.keys(vendaMap), ...Object.keys(ocupMap)])].sort().reverse();
  if (!meses.length) throw new Error('Nenhum PDF encontrado nas pastas VENDAS / OCUPAÇÃO.');

  // Baixa e parseia a planilha de eventos
  let eventosMap = {};
  try {
    const xlsxFile = await findEventosXlsx();
    if (xlsxFile) {
      const xlsxBuf = await downloadFile(xlsxFile.id);
      eventosMap = parseEventos(xlsxBuf);
    }
  } catch(e) { console.warn('[Eventos]', e.message); }

  const dados = {};
  for (const mes of meses) {
    const vf = vendaMap[mes];
    const of = ocupMap[mes];
    const [vBuf, oBuf] = await Promise.all([
      vf ? downloadFile(vf.id) : Promise.resolve(null),
      of ? downloadFile(of.id) : Promise.resolve(null)
    ]);
    const [vText, oText] = await Promise.all([
      vBuf ? pdfParse(vBuf).then(r => r.text) : Promise.resolve(''),
      oBuf ? pdfParse(oBuf).then(r => r.text) : Promise.resolve('')
    ]);
    const vendas   = parseVendas(vText);
    const ocupacao = parseOcupacao(oText);
    dados[mes] = buildMesData(vf, of, vendas, ocupacao, eventosMap[mes]);
  }

  const result = { sincAt: new Date().toISOString(), meses, dados };
  fs.writeFileSync(RESULT_FILE, JSON.stringify(result, null, 2));
  return result;
}

// ── Debug Eventos ─────────────────────────────────────────────────────────────
app.get('/api/debug-eventos', async (req, res) => {
  try {
    const xlsxFile = await findEventosXlsx();
    if (!xlsxFile) return res.json({ erro: 'Arquivo xlsx não encontrado na pasta.' });
    const buf = await downloadFile(xlsxFile.id);
    const wb  = XLSX.read(buf, { type: 'buffer' });
    const out  = { arquivo: xlsxFile.name, abas: [] };

    for (const sheetName of wb.SheetNames) {
      const key = SHEET_MES[sheetName.toUpperCase().trim().normalize('NFD').replace(/[̀-ͯ]/g,'')];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });

      let hRow = -1, cPax = -1, cBanq = -1, cForma = -1, cDataDbg = -1, headerCells = [];
      for (let i = 0; i < Math.min(15, rows.length); i++) {
        const r = rows[i].map(c => String(c).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,'').trim());
        const pi = r.findIndex(c => c.includes('PAX'));
        if (pi >= 0) {
          hRow = i; cPax = pi;
          cBanq    = r.findIndex(c => c.startsWith('BAN'));
          cForma   = r.findIndex(c => c.includes('FORMA') || (c.includes('PAGAMENTO') && c.length > 10));
          cDataDbg = r.findIndex(c => c === 'DATA' || c === 'DT' || c.startsWith('DATA'));
          headerCells = r;
          break;
        }
      }

      // Coleta todas as linhas não-vazias com PAX > 0 (para na primeira linha em branco)
      const linhasComputadas = [];
      let sumPax = 0, sumBanq = 0;
      for (let i = hRow + 1; i < rows.length; i++) {
        const row     = rows[i];
        const evento  = String(row[0] || '').trim();
        const paxRaw  = row[cPax];
        const banqRaw = row[cBanq];
        const forma   = String(row[cForma] || '').toUpperCase();
        if (!evento && (paxRaw === '' || paxRaw === null || paxRaw === undefined)) break;
        if (paxRaw === '' || paxRaw === null || paxRaw === undefined) continue;
        const pax = parseInt(paxRaw);
        if (!pax || isNaN(pax) || pax <= 0) continue;
        if (forma.includes('BACCO')) continue;
        const banq = typeof banqRaw === 'number'
          ? banqRaw
          : (parseFloat(String(banqRaw).replace(/[R$\s]/g,'').replace(/\./g,'').replace(',','.')) || 0);
        sumPax  += pax;
        sumBanq += banq;
        const dataRaw = cDataDbg >= 0 ? row[cDataDbg] : undefined;
        linhasComputadas.push({ idx: i, evento: row[0], pax, banq, forma: row[cForma], dataRaw, dataTipo: typeof dataRaw });
      }

      out.abas.push({
        nome: sheetName, mesKey: key || '(não mapeado)',
        headerLinha: hRow, headerCelulas: headerCells,
        colunas: { PAX: cPax, BAN: cBanq, FORMA: cForma, DATA: cDataDbg },
        totalLinhas: rows.length,
        RESULTADO: { totalPax: sumPax, totalBanq: +sumBanq.toFixed(2), linhasContadas: linhasComputadas.length },
        linhasComputadas
      });
    }
    res.json(out);
  } catch(e) { res.status(500).json({ erro: e.message }); }
});

// ── Debug arquivos ────────────────────────────────────────────────────────────
app.get('/api/debug-arquivos', async (req, res) => {
  try {
    const [vendaDir, ocupDir] = await Promise.all([
      findFolder(SHARED_DRIVE, 'VENDAS'),
      findFolder(SHARED_DRIVE, 'OCUPAÇÃO')
    ]);
    const [vendaPdfs, ocupPdfs] = await Promise.all([
      vendaDir ? allPdfs(vendaDir.id) : [],
      ocupDir  ? allPdfs(ocupDir.id)  : []
    ]);
    const xlsxFile = await findEventosXlsx().catch(() => null);
    let eventosAbas = [];
    if (xlsxFile) {
      const buf = await downloadFile(xlsxFile.id);
      const wb  = XLSX.read(buf, { type: 'buffer' });
      eventosAbas = wb.SheetNames.map(s => ({
        aba: s,
        mesKey: SHEET_MES[s.toUpperCase().trim().normalize('NFD').replace(/[̀-ͯ]/g,'')] || null
      }));
    }
    res.json({
      vendas:  vendaPdfs.map(f => ({ nome: f.name, mesDetectado: mesKey(f.name) })),
      ocupacao: ocupPdfs.map(f => ({ nome: f.name, mesDetectado: mesKey(f.name) })),
      eventos: { arquivo: xlsxFile?.name || null, abas: eventosAbas }
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Debug ─────────────────────────────────────────────────────────────────────
app.get('/api/debug-texto', async (req, res) => {
  try {
    const [vendaDir, ocupDir] = await Promise.all([
      findFolder(SHARED_DRIVE, 'VENDAS'),
      findFolder(SHARED_DRIVE, 'OCUPAÇÃO')
    ]);
    const [vendaPdf, ocupPdf] = await Promise.all([
      latestPdf(vendaDir.id),
      latestPdf(ocupDir.id)
    ]);
    const [vendaBuf, ocupBuf] = await Promise.all([
      downloadFile(vendaPdf.id),
      downloadFile(ocupPdf.id)
    ]);
    const [vendaText, ocupText] = await Promise.all([
      pdfParse(vendaBuf).then(r => r.text),
      pdfParse(ocupBuf).then(r => r.text)
    ]);
    const linhasTotal = vendaText.split('\n').filter(l => l.trim().startsWith('TOTAL'));
    res.json({
      vendas_inicio: vendaText.substring(0, 1500),
      vendas_fim: vendaText.substring(Math.max(0, vendaText.length - 2000)),
      linhas_TOTAL: linhasTotal.slice(-20),
      grand_parsed: parseVendas(vendaText).grand,
      ocupacao: ocupText.substring(0, 2000)
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Rotas ─────────────────────────────────────────────────────────────────────
app.get('/api/dados', (req, res) => {
  if (!fs.existsSync(RESULT_FILE)) return res.status(404).json({ error: 'Sem dados. Clique em Sincronizar.' });
  const store = JSON.parse(fs.readFileSync(RESULT_FILE, 'utf8'));
  // Legacy format (single month) — serve flat fields directly
  if (!store.meses) return res.json({ sincAt: store.sincAt, meses: [], mesSelecionado: null, mesLabel: 'Período atual', ...store });
  const mes = req.query.mes || store.meses[0];
  const d   = store.dados[mes];
  if (!d) return res.status(404).json({ error: `Mês ${mes} não encontrado.` });
  res.json({ sincAt: store.sincAt, meses: store.meses, mesSelecionado: mes, mesLabel: mesLabel(mes), ...d });
});

app.post('/api/sincronizar', async (req, res) => {
  try {
    const data = await sincronizar();
    res.json({ ok: true, data });
  } catch (e) {
    console.error('[Sync]', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/auth/google', (req, res) => {
  if (!CLIENT_ID) return res.status(500).send('GOOGLE_CLIENT_ID não configurado.');
  const p = new URLSearchParams({
    client_id: CLIENT_ID, redirect_uri: REDIRECT_URI,
    response_type: 'code', scope: 'https://www.googleapis.com/auth/drive.readonly',
    access_type: 'offline', prompt: 'consent'
  });
  res.redirect(`https://accounts.google.com/o/oauth2/v2/auth?${p}`);
});

app.get('/auth/callback', async (req2, res) => {
  const { code, error } = req2.query;
  if (error) return res.send(`<h2>Erro: ${error}</h2>`);
  try {
    const body = new URLSearchParams({
      client_id: CLIENT_ID, client_secret: CLIENT_SECRET,
      code, redirect_uri: REDIRECT_URI, grant_type: 'authorization_code'
    }).toString();
    const { status, body: rb } = await req('https://oauth2.googleapis.com/token',
      { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }, body);
    const tokens = JSON.parse(rb.toString());
    if (status !== 200) throw new Error(tokens.error_description || tokens.error);
    tokens.expiry = Date.now() + tokens.expires_in * 1000;
    saveTokens(tokens);
    res.send(`<html><body style="font-family:sans-serif;text-align:center;padding:60px">
      <h2 style="color:#1D9E75">✓ Google Drive autorizado!</h2>
      <p>Redirecionando...</p>
      <script>setTimeout(()=>location.href='/',1500)</script>
    </body></html>`);
  } catch (e) {
    res.status(500).send(`<h2>Erro: ${e.message}</h2>`);
  }
});

app.get('/auth/status', (req, res) => {
  const t = loadTokens();
  res.json({ autorizado: !!(t?.refresh_token), temClientId: !!CLIENT_ID, usuario: req.session?.user?.name || '' });
});

app.listen(PORT, '0.0.0.0', () => {
  console.log('\n========================================');
  console.log('  BACCO — Dashboard Gerencial');
  console.log(`  http://localhost:${PORT}`);
  console.log('========================================\n');
});
