process.env.TZ = 'America/Sao_Paulo';
// build-marker: forca novo deploy no Railway
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
const CUSTOS_CORRECOES_FILE = path.join(DATA_DIR, 'custos-correcoes.json');
const BEBIDAS_VINHOS_FILE = path.join(DATA_DIR, 'bebidas-vinhos.json');
const BEBIDAS_TACA_FILE = path.join(DATA_DIR, 'bebidas-vinhos-taca.json');
const BEBIDAS_MIGRACOES_FILE = path.join(DATA_DIR, 'bebidas-vinhos-migracoes.json');
const SHARED_DRIVE    = process.env.SHARED_DRIVE_ID    || '0AKZcsytstd78Uk9PVA';
const EVENTOS_FOLDER  = process.env.EVENTOS_FOLDER_ID  || '1OjS3q7vAccft_n4novmv6d86MBrwiQ9k';
const CAED_FILE_ID = process.env.CAED_FILE_ID || '1sRXE6m2UHVjC0oAjiYBydbsYzKrUmSQU7bmrjGDjkxg';
const CUSTOS_FILE_ID = process.env.CUSTOS_FILE_ID || '1DrhrWAqb3eIButKhj4J9HPydAadw7khG';
const TRIPADVISOR_DOC_ID = process.env.TRIPADVISOR_DOC_ID || '1yMNqdsXmD50fdLWVhSj4VII5vAFxqgFUHvoPfh5RVgg';
const CONSUMO_FOLDER_ID = process.env.CONSUMO_FOLDER_ID || '1gsvjga8clKukuN-S5AEWZHY6pHHT0RUn';
const CLIENT_ID    = process.env.GOOGLE_CLIENT_ID    || '';
const CLIENT_SECRET= process.env.GOOGLE_CLIENT_SECRET|| '';
const PORT         = process.env.PORT || 3001;
const BASE_URL     = process.env.BASE_URL || `http://localhost:${PORT}`;
const REDIRECT_URI = `${BASE_URL}/auth/callback`;
const SESSION_SECRET = process.env.SESSION_SECRET || 'bacco-gerencial-secret-2026';

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// ── Usuários ──────────────────────────────────────────────────────────────────
function hashPwd(pwd) { return crypto.createHash('sha256').update(pwd + 'bacco-salt').digest('hex'); }

const DEFAULT_USERS = {
  andrea: { displayName: 'Andrea', password: hashPwd('1234'),   mustChange: true },
  yoshio: { displayName: 'Yoshio', password: hashPwd('1234'),   mustChange: true },
  rafael: { displayName: 'Rafael', password: hashPwd('123456'), mustChange: true },
  paulo:  { displayName: 'Paulo',  password: hashPwd('123456'), mustChange: true }
};

function loadUsers() {
  if (!fs.existsSync(USERS_FILE)) {
    fs.writeFileSync(USERS_FILE, JSON.stringify(DEFAULT_USERS, null, 2));
    return DEFAULT_USERS;
  }
  const users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
  // Adiciona usuários padrão que ainda não existem (sem sobrescrever os já criados)
  let changed = false;
  for (const [id, u] of Object.entries(DEFAULT_USERS)) {
    if (!users[id]) { users[id] = u; changed = true; }
  }
  if (changed) fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2));
  return users;
}

function saveUsers(users) { fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2)); }

// ── Correções de Centro de Custo (aba Custos / Auditoria) ────────────────────
function loadCorrecoes() {
  try { return JSON.parse(fs.readFileSync(CUSTOS_CORRECOES_FILE, 'utf8')); }
  catch { return { auditoria: {}, reatribuidos: {} }; }
}
function saveCorrecoes(c) { fs.writeFileSync(CUSTOS_CORRECOES_FILE, JSON.stringify(c, null, 2)); }

// ── Bebidas / Vinhos — base inicial de estoque ────────────────────────────────
// Extrai o tamanho da embalagem (750ml, 375ml, 5lt...) do nome; usa a unidade (UN/GF) como fallback
function extraiTamanhoENome(nomeRaw, unidadeMedida) {
  const m = nomeRaw.match(/(\d+[.,]?\d*\s?(?:ml|lt|l)\b)/i);
  let tamanho = unidadeMedida;
  let nome = nomeRaw;
  if (m) {
    tamanho = m[1].trim();
    nome = (nomeRaw.slice(0, m.index) + nomeRaw.slice(m.index + m[0].length)).replace(/\s+/g, ' ').trim();
  }
  nome = nome.replace(/^vinhos?\s+|^vin\s+/i, '').trim();
  return { nome, tamanho };
}

const SEED_VINHOS_RAW = [
  ['100051','vin arg cordero cabernet tto 750ml','GF',1],
  ['000415','espum. casa geraldo natural brut','UN',3],
  ['001198','espumante alud branco 750ml','UN',1],
  ['1920','espumante chandon brut 750ml','GF',3],
  ['001202','vinho adele branco 750ml','UN',2],
  ['001203','vinho adele rose 750ml','UN',2],
  ['000935','vinho anubis malbec 750ml','UN',3],
  ['001199','vinho arcaia pinot grigio 750ml','UN',2],
  ['001208','vinho aresti sel chardonnay 750ml','UN',5],
  ['001204','vinho azul ventozelo','UN',2],
  ['000934','vinho bons ventos rose 375ml','UN',1],
  ['001200','vinho burdizzo primitivo 750ml','UN',3],
  ['001214','vinho cavas de oro blen rosado 750ml','UN',2],
  ['001213','vinho cavas de oro blend branco 750ml','UN',3],
  ['001201','vinho corsarini moltepulciano','UN',2],
  ['45645','vinho dom de minas cabernet franc 750ml','UN',1],
  ['20026','vinho e.a. fundacao eugenio de almeida','GF',3],
  ['000701','vinho folha do meio colheita branco','UN',5],
  ['000700','vinho folha do meio colheita tinto','UN',2],
  ['001211','vinho humberto can denario cab sauv 750m','UN',2],
  ['001212','vinho humberto canale den malbec 750ml','UN',1],
  ['001210','vinho los aljibes tinto 750ml','UN',3],
  ['000423','vinho luiz porto chardonnay 750ml','UN',10],
  ['001209','vinho spinoglio tierra alta tannat','UN',3],
  ['20081','vinho tinto bag 5 lt','UN',1],
  ['001216','vinho villa rosa colheita tinto','UN',2],
  ['000941','vinho white blend 3tons','UN',3]
];

// Vinhos que chegaram fora do estoque inicial contado — entram com Estoque Inicial 0;
// as quantidades são lançadas manualmente no campo Entradas (cumulativo)
const NOVOS_VINHOS_RAW = [
  ['N001', "ita caleo montepulciano d'abruzzo 750 ml", 'UN', 0],
  ['N002', 'codici primitivo puglia 750 ml', 'UN', 0],
  ['N003', 'le casine chianti 750 ml', 'UN', 0],
  ['N004', 'cartuxa ea tinto 750 ml', 'UN', 0],
  ['N005', 'tantehue carmenère 750 ml', 'UN', 0],
  ['N006', 'santa helena cabernet sauvignon 375 ml', 'UN', 0],
  ['N007', 'cartuxa ea alicante bouschet 750 ml', 'UN', 0],
  ['N008', 'cartuxa ea aragonez 750 ml', 'UN', 0],
  ['N009', 'cartuxa ea trincadeira 750 ml', 'UN', 0],
  ['06020043', 'por bons ventos tto 375ml', 'UN', 0],
  ['06020045', 'por bons ventos tto 750ml', 'UN', 0],
  ['VCC0003', 'arg cordero con piel de lobo cabernet sauvignon 750ml', 'UN', 0],
  ['14004', 'chi tarapaca reserva merlot 750ml', 'UN', 0],
  ['WC0109', 'arg cordero con piel de lobo malbec 750ml', 'UN', 0],
  // Vinhos encontrados na Contagem_Consolidada_Vinhos 28_07_2026.pdf que ainda não constavam na base
  ['G101', 'garibaldi primicias brut 660 ml', 'UN', 0],
  ['G102', 'garibaldi vero rose brut', 'UN', 0],
  ['G103', 'garibaldi vero pinot noir rose brut', 'UN', 0],
  ['G104', 'aresti reserva cabina 56 sauvignon blanc', 'UN', 0],
  ['G105', 'folha do meio rose', 'UN', 0],
  ['G106', 'anubis reserva malbec', 'UN', 0],
  ['G107', 'bons ventos magnum 1,5l', 'UN', 0],
  ['G108', 'ea cartuxa tinto 375 ml', 'UN', 0]
];

// Lançamentos de entrada iniciais que só podem ser aplicados uma vez (marca em BEBIDAS_MIGRACOES_FILE)
const ENTRADAS_INICIAIS_MIGRACAO = [
  { id: 'entrada-06020043-2026-07', codigo: '06020043', delta: 2 },
  { id: 'entrada-06020045-2026-07', codigo: '06020045', delta: 3 },
  { id: 'entrada-VCC0003-2026-07',  codigo: 'VCC0003',  delta: 3 },
  { id: 'entrada-14004-2026-07',    codigo: '14004',    delta: 3 },
  { id: 'entrada-WC0109-2026-07',   codigo: 'WC0109',   delta: 3 },
  { id: 'entrada-N005-200244-2026-07', codigo: 'N005',  delta: 3 } // Tantehue Carmenère já cadastrado
];

// Contagem física de 28/07/2026 (Contagem_Consolidada_Vinhos.pdf) — aplicada uma única vez no
// campo Auditoria; o usuário pode sobrescrever livremente depois com a contagem real do dia 31
const AUDITORIA_28_07_2026 = [
  { codigo: '000415', valor: 3 },   // Casa Geraldo Glera Brut
  { codigo: '001198', valor: 1 },   // Alud Branco
  { codigo: '1920',   valor: 3 },   // Chandon Brut
  { codigo: '000941', valor: 3 },   // Casa Geraldo White Blend 3 Tons
  { codigo: '001208', valor: 3 },   // Aresti Estate Select Chardonnay
  { codigo: '000701', valor: 5 },   // Folha do Meio Colheita (branco)
  { codigo: '001199', valor: 1 },   // Arcaia Pinot Grigio
  { codigo: '001213', valor: 3 },   // Cavas de Oro Blend Branco
  { codigo: '000423', valor: 8 },   // Luiz Porto Chardonnay
  { codigo: '001203', valor: 2 },   // Adele (rosé)
  { codigo: '001214', valor: 2 },   // Cavas de Oro Blend (rosé)
  { codigo: '001210', valor: 3 },   // Los Aljibes (VA) Tempranillo
  { codigo: '001209', valor: 2 },   // Spinoglio Tierra Alta Tannat
  { codigo: '001201', valor: 0 },   // Corsarini Montepulciano d'Abruzzo
  { codigo: '001200', valor: 0 },   // Burdizzo Primitivo
  { codigo: '000935', valor: 3 },   // Anubis Malbec
  { codigo: 'WC0109', valor: 0 },   // Cordero con Piel de Lobo Malbec
  { codigo: 'VCC0003',valor: 3 },   // Cordero con Piel de Lobo Cabernet Sauvignon
  { codigo: '000700', valor: 2 },   // Folha do Meio Colheita (tinto)
  { codigo: '06020043', valor: 0 },// Bons Ventos 375ml
  { codigo: '06020045', valor: 4 },// Bons Ventos
  { codigo: 'N004',   valor: 3 },  // EA Tinto
  { codigo: 'N007',   valor: 2 },  // EA Alicante Bouschet
  { codigo: 'N008',   valor: 2 },  // EA Aragonez
  { codigo: 'N009',   valor: 1 },  // EA Trincadeira
  { codigo: 'N006',   valor: 2 },  // Santa Helena Cabernet Sauvignon 375ml
  { codigo: 'N005',   valor: 0 },  // Tantahue Carmenere
  { codigo: 'N001',   valor: 3 },  // Caleo Montepulciano d'Abruzzo
  { codigo: 'N002',   valor: 2 },  // Codici Primitivo Puglia
  { codigo: 'N003',   valor: 2 },  // Le Cascine Chianti
  { codigo: '14004',  valor: 2 },  // Tarapacá
  { codigo: '001204', valor: 2 },  // Azul de Ventos 375ml (casado com Azul Ventozelo — confira)
  { codigo: '001216', valor: 1 },  // Villa Rosa
  { codigo: 'G101', valor: 10 },   // Garibaldi Primícias Brut
  { codigo: 'G102', valor: 4 },    // Garibaldi Vero Rosé Brut
  { codigo: 'G103', valor: 5 },    // Garibaldi Vero Pinot Noir Rosé Brut
  { codigo: 'G104', valor: 2 },    // Aresti Reserva Cabina 56 Sauvignon Blanc
  { codigo: 'G105', valor: 1 },    // Folha do Meio (rosé)
  { codigo: 'G106', valor: 0 },    // Anubis Reserva Malbec
  { codigo: 'G107', valor: 4 },    // Bons Ventos Magnum 1,5L
  { codigo: 'G108', valor: 0 }     // EA Cartuxa Tinto 375ml
];

// Categoria inferida pelo nome/uva de cada rótulo (pesquisada onde disponível).
// Deixado em branco quando não há certeza — ajuste manual necessário.
const CATEGORIA_POR_CODIGO = {
  '100051': 'Tinto',        // Cabernet
  '000415': 'Espumante',
  '001198': 'Espumante',
  '1920':   'Espumante',
  '001202': 'Branco',
  '001203': 'Rosé',
  '000935': 'Tinto',        // Malbec
  '001199': 'Branco',       // Pinot Grigio
  '001208': 'Branco',       // Chardonnay
  '001204': '',             // Azul Ventozelo — não confirmado
  '000934': 'Rosé',
  '001200': 'Tinto',        // Primitivo
  '001214': 'Rosé',
  '001213': 'Branco',
  '001201': 'Tinto',        // Montepulciano
  '45645':  'Tinto',        // Cabernet Franc
  '20026':  '',             // EA Fundação Eugénio de Almeida — não confirmado
  '000701': 'Branco',
  '000700': 'Tinto',
  '001211': 'Tinto',        // Cabernet Sauvignon
  '001212': 'Tinto',        // Malbec
  '001210': 'Tinto',
  '000423': 'Branco',       // Chardonnay
  '001209': 'Tinto',        // Tannat
  '20081':  'Tinto',
  '001216': 'Tinto',
  '000941': 'Branco',       // White Blend
  'N001':   'Tinto',        // Montepulciano d'Abruzzo
  'N002':   'Tinto',        // Primitivo
  'N003':   'Tinto',        // Chianti
  'N004':   'Tinto',
  'N005':   'Tinto',        // Carménère
  'N006':   'Tinto',        // Cabernet Sauvignon
  'N007':   'Tinto',        // Alicante Bouschet
  'N008':   'Tinto',        // Aragonez
  'N009':   'Tinto',        // Trincadeira
  '06020043': 'Tinto',
  '06020045': 'Tinto',
  'VCC0003':  'Tinto',      // Cabernet Sauvignon
  '14004':    'Tinto',      // Merlot
  'WC0109':   'Tinto',      // Malbec
  'G101': 'Espumante', // Garibaldi Primícias Brut
  'G102': 'Espumante', // Garibaldi Vero Rosé Brut
  'G103': 'Espumante', // Garibaldi Vero Pinot Noir Rosé Brut
  'G104': 'Branco',    // Aresti Reserva Cabina 56 Sauvignon Blanc
  'G105': 'Rosé',      // Folha do Meio Rosé
  'G106': 'Tinto',     // Anubis Reserva Malbec
  'G107': 'Tinto',     // Bons Ventos Magnum 1,5L
  'G108': 'Tinto'      // EA Cartuxa Tinto 375ml
};
const ORDEM_CATEGORIAS = ['Espumante', 'Branco', 'Rosé', 'Tinto'];

function seedItensDe(lista) {
  return lista.map(([codigo, nomeRaw, unidade, estoqueInicial]) => {
    const { nome, tamanho } = extraiTamanhoENome(nomeRaw, unidade);
    return { codigo, nome, tamanho, categoria: CATEGORIA_POR_CODIGO[codigo] || '', estoqueInicial, vendas: 0, entradas: 0, auditoria: null, observacao: '' };
  });
}
function seedVinhos() { return seedItensDe(SEED_VINHOS_RAW); }

function loadVinhos() {
  const existeArquivo = fs.existsSync(BEBIDAS_VINHOS_FILE);
  let itens = existeArquivo ? JSON.parse(fs.readFileSync(BEBIDAS_VINHOS_FILE, 'utf8')) : seedItensDe(SEED_VINHOS_RAW);

  // Mescla vinhos novos que ainda não estão na base salva (sem sobrescrever os já existentes/editados)
  const codigosExistentes = new Set(itens.map(i => i.codigo));
  let mudou = !existeArquivo;
  for (const novo of seedItensDe(NOVOS_VINHOS_RAW)) {
    if (!codigosExistentes.has(novo.codigo)) { itens.push(novo); mudou = true; }
  }

  // Preenche a categoria uma única vez para itens que ainda não têm esse campo
  // (não sobrescreve se o usuário já editou manualmente, mesmo que tenha deixado em branco)
  for (const item of itens) {
    if (item.categoria === undefined) { item.categoria = CATEGORIA_POR_CODIGO[item.codigo] || ''; mudou = true; }
  }

  // Aplica lançamentos de entrada iniciais uma única vez
  const migracoes = fs.existsSync(BEBIDAS_MIGRACOES_FILE)
    ? JSON.parse(fs.readFileSync(BEBIDAS_MIGRACOES_FILE, 'utf8'))
    : { aplicadas: [] };
  let migracoesMudaram = false;
  for (const m of ENTRADAS_INICIAIS_MIGRACAO) {
    if (migracoes.aplicadas.includes(m.id)) continue;
    const item = itens.find(i => i.codigo === m.codigo);
    if (item) { item.entradas = +((item.entradas || 0) + m.delta).toFixed(3); mudou = true; }
    migracoes.aplicadas.push(m.id);
    migracoesMudaram = true;
  }

  // Aplica a contagem física de 28/07/2026 no campo Auditoria — só uma vez (o usuário pode
  // sobrescrever livremente depois com a contagem do dia 31)
  for (const a of AUDITORIA_28_07_2026) {
    const id = `auditoria-28-07-2026-${a.codigo}`;
    if (migracoes.aplicadas.includes(id)) continue;
    const item = itens.find(i => i.codigo === a.codigo);
    if (item) { item.auditoria = a.valor; mudou = true; }
    migracoes.aplicadas.push(id);
    migracoesMudaram = true;
  }
  if (migracoesMudaram) fs.writeFileSync(BEBIDAS_MIGRACOES_FILE, JSON.stringify(migracoes, null, 2));

  if (mudou) fs.writeFileSync(BEBIDAS_VINHOS_FILE, JSON.stringify(itens, null, 2));
  return itens;
}
function saveVinhos(itens) { fs.writeFileSync(BEBIDAS_VINHOS_FILE, JSON.stringify(itens, null, 2)); }

// ── Sessão e middleware ───────────────────────────────────────────────────────
app.use(session({ secret: SESSION_SECRET, resave: false, saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 } }));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));

const PUBLIC_PATHS = ['/login', '/login.html', '/trocar-senha', '/trocar-senha.html'];
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
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

app.post('/api/admin/reset-senha', (req, res) => {
  if (!req.session?.user) return res.status(401).json({ error: 'Não autenticado.' });
  const { usuario, novaSenha } = req.body;
  const id = String(usuario || '').toLowerCase();
  const users = loadUsers();
  if (!users[id]) return res.status(404).json({ error: `Usuário ${id} não encontrado.` });
  users[id].password   = hashPwd(novaSenha || '123456');
  users[id].mustChange = true;
  saveUsers(users);
  res.json({ ok: true, usuario: id, mustChange: true });
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
  if (status === 403 || status === 400) {
    // Provavelmente é uma planilha nativa do Google Sheets (não um .xlsx real) — precisa exportar em vez de baixar
    const token2 = await getToken();
    const exportMime = encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const exportRes = await req(
      `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=${exportMime}`,
      { headers: { Authorization: `Bearer ${token2}` } }
    );
    if (exportRes.status !== 200) throw new Error(`Download falhou HTTP ${status} (export também falhou: ${exportRes.status})`);
    return exportRes.body;
  }
  if (status !== 200) throw new Error(`Download falhou HTTP ${status}`);
  return body;
}

async function downloadGoogleDocText(fileId) {
  const token = await getToken();
  const { status, body } = await req(
    `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=text/plain`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (status !== 200) throw new Error(`Export do Google Doc falhou HTTP ${status}`);
  return body.toString('utf8');
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
      const nums = [...line.matchAll(/\d{1,3}(?:\.\d{3})*,\d{2}/g)].map(m => parseFloat(m[0].replace(/\./g,'').replace(',','.')));
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
      const nums = [...bloco.matchAll(/\d{1,3}(?:\.\d{3})*,\d{2}/g)].map(m => parseFloat(m[0].replace(/\./g,'').replace(',','.')));
      // As últimas 5 colunas são sempre Bruto|Desconto|Líquido|Taxa|Total Pago,
      // independente de quantas colunas variáveis (QTD, CUSTO) vierem antes
      if (nums.length >= 5) {
        const [valorBruto, desconto, valorLiquido, taxaServico, totalPago] = nums.slice(-5);
        grand = { valorBruto, desconto, valorLiquido, taxaServico, totalPago };
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

    const [mesAno, mesMes] = mesKey.split('-');

    // Converte célula DATA para dd/mm/yyyy — suporta serial Excel, ranges, mês implícito
    const parseDateCell = v => {
      if (v === null || v === undefined || v === '') return null;

      // Serial Excel (inteiro ou decimal como 46003.33)
      if (typeof v === 'number') {
        const serial = Math.round(v);
        const dt = new Date((serial - 25569) * 86400 * 1000);
        const dd = String(dt.getUTCDate()).padStart(2,'0');
        const mm = String(dt.getUTCMonth()+1).padStart(2,'0');
        return `${dd}/${mm}/${dt.getUTCFullYear()}`;
      }

      const s = String(v).trim();
      if (!s) return null;

      // Caso 1: começa com dd/mm ou dd/mm/yyyy — pega a 1ª data (ignora range seguinte)
      const m1 = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/);
      if (m1) {
        const dd = m1[1].padStart(2,'0');
        const mm = m1[2].padStart(2,'0');
        const yr = m1[3] ? (m1[3].length === 2 ? '20'+m1[3] : m1[3]) : mesAno;
        return `${dd}/${mm}/${yr}`;
      }

      // Caso 2: "dd E/A dd/mm" — 1º dia + mês no fim do range (ex: "06 E 07/01", "29 e30/04")
      const m2 = s.match(/^(\d{1,2})\s*[EeAa]\s*\d{1,2}\/(\d{1,2})/);
      if (m2) {
        return `${m2[1].padStart(2,'0')}/${m2[2].padStart(2,'0')}/${mesAno}`;
      }

      // Caso 3: "dd E/A dd" sem mês — usa mês da aba (ex: "29 E 30", "12 A 20")
      const m3 = s.match(/^(\d{1,2})\s*[EeAa]\s*\d{1,2}(\s|$)/);
      if (m3) {
        return `${m3[1].padStart(2,'0')}/${mesMes}/${mesAno}`;
      }

      // Caso 4: só "dd" — usa mês e ano da aba
      const m4 = s.match(/^(\d{1,2})$/);
      if (m4) {
        return `${m4[1].padStart(2,'0')}/${mesMes}/${mesAno}`;
      }

      return null;
    };

    let totalPax = 0, totalBanq = 0, totalSala = 0, totalEquip = 0;
    const daily = {};
    const linhas = [];
    let linhasVazias = 0;
    for (let i = hRow + 1; i < rows.length; i++) {
      const row     = rows[i];
      const evento  = String(row[0] || '').trim();
      const paxRaw  = row[cPax];
      const banqRaw = row[cBanq];
      const forma   = String(row[cForma] || '').toUpperCase();

      // Conta linhas consecutivas sem evento E sem PAX — para após 3
      const vazia = !evento && (paxRaw === '' || paxRaw === null || paxRaw === undefined);
      if (vazia) { if (++linhasVazias >= 3) break; continue; }
      linhasVazias = 0;

      if (paxRaw === '' || paxRaw === null || paxRaw === undefined) continue;
      const pax = parseInt(paxRaw);
      if (!pax || isNaN(pax) || pax <= 0) continue;
      if (forma.includes('BACCO')) continue;

      const banq  = parseVal(banqRaw);
      const sala  = cSala  >= 0 ? parseVal(row[cSala])  : 0;
      const equip = cEquip >= 0 ? parseVal(row[cEquip]) : 0;
      const rowTotal = +(sala + equip + banq).toFixed(2);
      const dateKey = cData >= 0 ? parseDateCell(row[cData]) : null;

      totalPax  += pax;
      totalBanq += banq;
      totalSala += sala;
      totalEquip+= equip;

      // Acumula por data se coluna DATA disponível
      if (dateKey) {
        if (!daily[dateKey]) daily[dateKey] = { pax:0, sala:0, equip:0, banq:0, total:0 };
        daily[dateKey].pax   += pax;
        daily[dateKey].sala  += sala;
        daily[dateKey].equip += equip;
        daily[dateKey].banq  += banq;
        daily[dateKey].total += rowTotal;
      }

      linhas.push({ data: dateKey || '', evento, pax, banq: +banq.toFixed(2), sala: +sala.toFixed(2), equip: +equip.toFixed(2), total: rowTotal, forma: String(row[cForma] || '').trim() });
    }
    const total = +(totalSala + totalEquip + totalBanq).toFixed(2);
    result[mesKey] = { pax: totalPax, sala: +totalSala.toFixed(2), equip: +totalEquip.toFixed(2), banq: +totalBanq.toFixed(2), total, daily, linhas };
  }
  return result;
}

// ── Parser CAEd (check-in/check-out → diárias dia a dia) ─────────────────────
const CAED_ANO_PADRAO = 2026;
const CAED_RATE_MEDIA = 46; // taxa única por diária (R$), independente do tipo de pensão

function normalizaTexto(s) {
  return String(s || '').trim().toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function serialParaData(serial) {
  const dt = new Date((serial - 25569) * 86400 * 1000);
  return new Date(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate());
}

function parseDataDDMM(v, anoDefault) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return serialParaData(Math.round(v));

  const s = String(v).trim();
  // Célula com data "real" do Excel/Sheets renderizada como número puro (raw:false não formatou como data)
  if (/^\d{4,6}(\.\d+)?$/.test(s)) return serialParaData(Math.round(parseFloat(s)));

  const m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (!m) return null;
  const dd = +m[1], mm = +m[2];
  const yr = m[3] ? (m[3].length === 2 ? 2000 + +m[3] : +m[3]) : anoDefault;
  return new Date(yr, mm - 1, dd);
}

// Preenche células mescladas com o valor da célula superior-esquerda, para que
// reservas com Check-in/Check-out mesclados em várias linhas (um hóspede por linha) sejam lidas corretamente.
function expandirMescladas(sheet) {
  const merges = sheet['!merges'] || [];
  for (const range of merges) {
    const topLeftAddr = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
    const topLeftCell = sheet[topLeftAddr];
    if (!topLeftCell) continue;
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        if (r === range.s.r && c === range.s.c) continue;
        const addr = XLSX.utils.encode_cell({ r, c });
        sheet[addr] = { ...topLeftCell };
      }
    }
  }
}

const TIPO_LABEL = { 'SEM PENSAO': 'Sem Pensão', 'MEIA PENSAO': 'Meia Pensão', 'PENSAO COMPLETA': 'Pensão Completa' };

function parseCaed(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const porMes = {}; // { 'YYYY-MM': { daily: {...}, linhas: [...] } }

  const getMes = mesKeyStr => {
    if (!porMes[mesKeyStr]) porMes[mesKeyStr] = { daily: {}, linhas: [] };
    return porMes[mesKeyStr];
  };

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    expandirMescladas(sheet);
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });

    for (const row of rows) {
      const col0 = normalizaTexto(row[0]);
      if (col0 === 'RESERVA') continue; // linha de cabeçalho (pode se repetir várias vezes na planilha)

      const nomeCel = String(row[1] || '').trim();
      if (!nomeCel) continue;
      const checkin  = parseDataDDMM(row[2], CAED_ANO_PADRAO);
      const checkout = parseDataDDMM(row[3], CAED_ANO_PADRAO);
      if (!checkin || !checkout || checkout <= checkin) continue;

      const tipoPensao = normalizaTexto(row[5]);
      const rate = CAED_RATE_MEDIA;
      const status = String(row[6] || '').trim();
      const nomes = nomeCel.split('\n').map(n => n.trim()).filter(Boolean);
      // Diária é contada por RESERVA (quarto), não por hóspede — reservas com 2+ nomes
      // compartilham o mesmo quarto/diária, então não deve multiplicar pela qtd de pessoas
      const totalNoites = Math.round((checkout - checkin) / 86400000);

      // Uma linha por hóspede na tabela da aba CAEd, no mês do check-in
      const dd0 = String(checkin.getDate()).padStart(2, '0');
      const mm0 = String(checkin.getMonth() + 1).padStart(2, '0');
      const yyyy0 = checkin.getFullYear();
      const dd1 = String(checkout.getDate()).padStart(2, '0');
      const mm1 = String(checkout.getMonth() + 1).padStart(2, '0');
      const yyyy1 = checkout.getFullYear();
      const mesCheckin = getMes(`${yyyy0}-${mm0}`);
      for (const nome of nomes) {
        mesCheckin.linhas.push({
          nome, checkin: `${dd0}/${mm0}/${yyyy0}`, checkout: `${dd1}/${mm1}/${yyyy1}`,
          diarias: totalNoites, tipoPensao: TIPO_LABEL[tipoPensao] || (row[5] || '—'),
          valor: +(totalNoites * rate).toFixed(2), status
        });
      }

      // Cada noite de [checkin, checkout) conta como uma diária (checkout não é contado)
      for (let d = new Date(checkin); d < checkout; d.setDate(d.getDate() + 1)) {
        const dd = String(d.getDate()).padStart(2, '0');
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const yyyy = d.getFullYear();
        const mesAtual = getMes(`${yyyy}-${mm}`);
        const dataStr = `${dd}/${mm}/${yyyy}`;
        if (!mesAtual.daily[dataStr]) mesAtual.daily[dataStr] = { pessoas: 0, faturamento: 0, semPensao: 0, meiaPensao: 0, pensaoCompleta: 0 };
        const dia = mesAtual.daily[dataStr];
        dia.pessoas += 1;
        dia.faturamento += rate;
        if (tipoPensao === 'SEM PENSAO') dia.semPensao += 1;
        else if (tipoPensao === 'MEIA PENSAO') dia.meiaPensao += 1;
        else if (tipoPensao === 'PENSAO COMPLETA') dia.pensaoCompleta += 1;
      }
    }
  }

  const result = {};
  for (const [mesKeyStr, { daily, linhas }] of Object.entries(porMes)) {
    let totalPessoas = 0, totalFaturamento = 0, semPensao = 0, meiaPensao = 0, pensaoCompleta = 0;
    for (const v of Object.values(daily)) {
      v.faturamento = +v.faturamento.toFixed(2);
      totalPessoas += v.pessoas;
      totalFaturamento += v.faturamento;
      semPensao += v.semPensao; meiaPensao += v.meiaPensao; pensaoCompleta += v.pensaoCompleta;
    }
    linhas.sort((a, b) => a.checkin.split('/').reverse().join('').localeCompare(b.checkin.split('/').reverse().join('')));
    result[mesKeyStr] = {
      daily, linhas, totalPessoas, semPensao, meiaPensao, pensaoCompleta,
      totalFaturamento: +totalFaturamento.toFixed(2)
    };
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
function buildMesData(vendaPdf, ocupPdf, vendas, ocupacao, eventosmesRaw, mes, caedMes) {
  // No mês corrente, eventos agendados para datas futuras não devem compor o faturamento realizado
  let eventosmes = eventosmesRaw;
  if (eventosmesRaw && mes === mesAtualKey()) {
    const hoje = new Date(); hoje.setHours(23,59,59,999);
    const linhasPassadas = (eventosmesRaw.linhas || []).filter(l => {
      if (!l.data) return false;
      const [dd,mm,yyyy] = l.data.split('/');
      return new Date(+yyyy, +mm-1, +dd) <= hoje;
    });
    const soma = linhasPassadas.reduce((a,l)=>({
      pax: a.pax+(l.pax||0), banq: a.banq+(l.banq||0), sala: a.sala+(l.sala||0),
      equip: a.equip+(l.equip||0), total: a.total+(l.total||0)
    }), {pax:0,banq:0,sala:0,equip:0,total:0});
    const dailyPassado = {};
    for (const [data, v] of Object.entries(eventosmesRaw.daily || {})) {
      const [dd,mm,yyyy] = data.split('/');
      if (new Date(+yyyy, +mm-1, +dd) <= hoje) dailyPassado[data] = v;
    }
    eventosmes = { pax: soma.pax, banq: +soma.banq.toFixed(2), sala: +soma.sala.toFixed(2),
      equip: +soma.equip.toFixed(2), total: +soma.total.toFixed(2), daily: dailyPassado, linhas: linhasPassadas };
  }
  const totalRSTdiario = Object.values(vendas.daily).reduce((s,d) => s + (d.RESTAURANTE||0), 0);
  const totalRSdiario  = Object.values(vendas.daily).reduce((s,d) => s + (d['Room Service']||0), 0);
  // Usa bruto+taxa por PDV (acumulado dos TOTAL DO DIA), com fallback nos totais diários
  const totalRST  = vendas.brutoTaxa?.RESTAURANTE  || totalRSTdiario;
  const totalRS   = vendas.brutoTaxa?.['Room Service'] || totalRSdiario;
  const totalPago = +(totalRST + totalRS).toFixed(2);
  const clientes  = vendas.notas.RESTAURANTE + vendas.notas['Room Service'];
  const hospedes  = ocupacao.total;

  const allDates = [...new Set([...Object.keys(vendas.daily), ...Object.keys(ocupacao.daily)])].sort();

  // Faturamento de Eventos considera apenas Banquete (não soma Sala/Equipamentos)
  const eventosBanqPre   = eventosmes?.banq  || 0;

  // Usa breakdown diário de eventos se disponível; caso contrário distribui igualmente
  const eventosDailyMap = eventosmes?.daily || {};
  const temDailyEventos = Object.keys(eventosDailyMap).length > 0;
  const eventosDiarioFallback = (!temDailyEventos && allDates.length > 0)
    ? +(eventosBanqPre / allDates.length).toFixed(2) : 0;

  const serie = allDates.map(d => ({
    data:        d,
    restaurante: vendas.daily[d]?.RESTAURANTE || 0,
    roomService: vendas.daily[d]?.['Room Service'] || 0,
    cafe:        ocupacao.daily[d]?.receita  || 0,
    eventos:     temDailyEventos ? (eventosDailyMap[d]?.banq || 0) : eventosDiarioFallback,
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
  const receitaEventos  = eventosBanq;

  // CAEd no Dashboard conta do dia 1 até a data de hoje (independente do PDF de vendas,
  // já que o CAEd vem de uma planilha própria no Drive)
  let clientesCaed = 0, receitaCaed = 0;
  if (caedMes?.daily) {
    const hoje = new Date(); hoje.setHours(23, 59, 59, 999);
    for (const [dataStr, v] of Object.entries(caedMes.daily)) {
      const [dd, mm, yyyy] = dataStr.split('/');
      const dataAtual = new Date(+yyyy, +mm - 1, +dd);
      if (dataAtual > hoje) continue;
      clientesCaed += v.pessoas;
      receitaCaed  += v.faturamento;
    }
  }
  receitaCaed = +receitaCaed.toFixed(2);
  const totalGeral      = +(totalPago + receitaCafe + receitaEventos + receitaCaed).toFixed(2);
  const totalClientes   = clientes + clientesCafe + clientesEventos + clientesCaed;
  const ticketCafe      = clientesCafe    > 0 ? Math.round(receitaCafe    / clientesCafe)             : 0;
  const ticketEventos   = clientesEventos > 0 ? +(receitaEventos / clientesEventos).toFixed(2)       : 0;
  const ticketCaed      = clientesCaed    > 0 ? +(receitaCaed    / clientesCaed   ).toFixed(2)       : 0;
  const ticketGeral     = totalClientes   > 0 ? +(totalGeral     / totalClientes  ).toFixed(2)       : 0;
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
    receitaCaed,
    faturamentoTotal:     totalGeral,
    // Clientes por canal
    clientesBacco:        vendas.notas.RESTAURANTE,
    clientesRoomService:  vendas.notas['Room Service'],
    clientesCafe,
    clientesEventos,
    clientesCaed,
    clientesTotal:        totalClientes,
    // Tickets por canal
    ticketRST:            vendas.notas.RESTAURANTE > 0 ? +(totalRST/vendas.notas.RESTAURANTE).toFixed(2) : 0,
    ticketRS:             vendas.notas['Room Service'] > 0 ? +(totalRS/vendas.notas['Room Service']).toFixed(2) : 0,
    ticketCafe,
    ticketEventos,
    ticketCaed,
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

  // Group by mes key derived from filename — em caso de duplicidade, mantém o mais recentemente modificado
  const vendaMap = {};
  for (const f of vendaPdfs) {
    const k = mesKey(f.name);
    if (k && (!vendaMap[k] || f.modifiedTime > vendaMap[k].modifiedTime)) vendaMap[k] = f;
  }
  const ocupMap  = {};
  for (const f of ocupPdfs) {
    const k = mesKey(f.name);
    if (k && (!ocupMap[k] || f.modifiedTime > ocupMap[k].modifiedTime)) ocupMap[k] = f;
  }

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

  // Baixa e parseia a planilha do CAEd (check-in/check-out)
  let caedMap = {};
  try {
    const caedBuf = await downloadFile(CAED_FILE_ID);
    caedMap = parseCaed(caedBuf);
  } catch(e) { console.warn('[CAEd]', e.message); }

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
    dados[mes] = buildMesData(vf, of, vendas, ocupacao, eventosMap[mes], mes, caedMap[mes]);

    // Apura vendas de vinho do mês corrente (ciclo mensal — recalcula do zero a cada sincronização)
    if (mes === mesAtualKey() && vText) {
      try { apurarVendasVinhosDoMes(vText); } catch(e) { console.warn('[Bebidas/Vinhos]', e.message); }
    }
  }

  const result = { sincAt: new Date().toISOString(), meses, dados, eventosRaw: eventosMap, caedRaw: caedMap };
  fs.writeFileSync(RESULT_FILE, JSON.stringify(result, null, 2));
  return result;
}

// ── Debug Inventário ──────────────────────────────────────────────────────────
app.get('/api/debug-caed', async (req, res) => {
  try {
    const buf = await downloadFile(CAED_FILE_ID);
    const resultado = parseCaed(buf);
    const mes = req.query.mes || Object.keys(resultado).sort().reverse()[0];
    res.json({ mesesEncontrados: Object.keys(resultado).sort(), mesExibido: mes, dados: resultado[mes] || null });
  } catch(e) { res.status(500).json({ erro: e.message }); }
});

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

      // Coleta todas as linhas não-vazias com PAX > 0 (para após 3 linhas em branco consecutivas)
      const linhasComputadas = [];
      let sumPax = 0, sumBanq = 0, vazias = 0;
      for (let i = hRow + 1; i < rows.length; i++) {
        const row     = rows[i];
        const evento  = String(row[0] || '').trim();
        const paxRaw  = row[cPax];
        const banqRaw = row[cBanq];
        const forma   = String(row[cForma] || '').toUpperCase();
        const vazia = !evento && (paxRaw === '' || paxRaw === null || paxRaw === undefined);
        if (vazia) { if (++vazias >= 3) break; continue; }
        vazias = 0;
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
function parseConsumoFilename(name) {
  // Espera nomes como "GASMIG 072026.PDF" → mm=07, yyyy=2026
  const m = String(name).match(/(\d{2})\s*(\d{4})/);
  if (!m) return null;
  const mm = +m[1], yyyy = +m[2];
  if (mm < 1 || mm > 12) return null;
  return `${yyyy}-${String(mm).padStart(2,'0')}`;
}

async function listaArquivosConsumo(prefixo) {
  const q = encodeURIComponent(`'${CONSUMO_FOLDER_ID}' in parents and trashed=false and mimeType='application/pdf'`);
  const d = await driveGet(`files?q=${q}&fields=files(id,name,modifiedTime)&corpora=allDrives`);
  const arquivos = (d.files || [])
    .filter(f => normalizaTexto(f.name).startsWith(prefixo))
    .map(f => ({ ...f, mesKey: parseConsumoFilename(f.name) }))
    .filter(f => f.mesKey);
  arquivos.sort((a, b) => b.mesKey.localeCompare(a.mesKey));
  return arquivos;
}

async function findArquivoConsumoMaisRecente(prefixo) {
  const arquivos = await listaArquivosConsumo(prefixo);
  return arquivos[0] || null;
}

// ── Parser GASMIG (consumo de gás por período de leitura) ────────────────────
function parseGasmigTexto(texto) {
  const numRe = /\d{1,3}(?:\.\d{3})*,\d{2,4}/g;

  // Período de faturamento atual (a linha "Período" seguida de "dd/mm/aa a dd/mm/aa")
  const atualMatch = texto.match(/Per[ií]odo\s*\n(\d{2})\/(\d{2})\/(\d{2})\s*a\s*(\d{2})\/(\d{2})\/(\d{2})/);
  let periodoAtual = null;
  if (atualMatch) {
    const inicio = new Date(2000 + (+atualMatch[3]), +atualMatch[2] - 1, +atualMatch[1]);
    const fim    = new Date(2000 + (+atualMatch[6]), +atualMatch[5] - 1, +atualMatch[4]);
    const medIdx = texto.indexOf('DADOS DE MEDIÇÃO');
    const fatIdx = texto.indexOf('DADOS DE FATURAMENTO');
    let consumo = null;
    if (medIdx >= 0 && fatIdx > medIdx) {
      const trecho = texto.substring(medIdx, fatIdx);
      const nums = [...trecho.matchAll(numRe)].map(m => parseFloat(m[0].replace(/\./g,'').replace(',','.')));
      if (nums.length) consumo = nums[nums.length - 1];
    }
    let totalPagar = null;
    const totalMatch = texto.match(/Total a pagar\s*\nR\$\s*([\d.,]+)/);
    if (totalMatch) totalPagar = parseFloat(totalMatch[1].replace(/\./g,'').replace(',','.'));

    if (consumo !== null) periodoAtual = { inicio, fim, consumo, totalPagar };
  }

  // Tabela HISTÓRICO (períodos anteriores)
  const historico = [];
  const histIdx = texto.indexOf('HISTÓRICO');
  if (histIdx >= 0) {
    const trecho = texto.substring(histIdx);
    const re = /(\d{2})\/(\d{2})\/(\d{4})\s*a\s*(\d{2})\/(\d{2})\/(\d{4})\s+([\d.,]+)/g;
    let m;
    while ((m = re.exec(trecho))) {
      historico.push({
        inicio: new Date(+m[3], +m[2] - 1, +m[1]),
        fim:    new Date(+m[6], +m[5] - 1, +m[4]),
        consumo: parseFloat(m[7].replace(/\./g,'').replace(',','.'))
      });
    }
  }

  const periodos = [];
  if (periodoAtual) periodos.push(periodoAtual);
  periodos.push(...historico);
  return periodos;
}

// Soma hóspedes (ocupação) de todos os meses já sincronizados dentro de um intervalo de datas
function somaOcupacaoPeriodo(store, inicio, fim) {
  let total = 0, temDados = false;
  for (const mesKeyStr of Object.keys(store.dados || {})) {
    const serie = store.dados[mesKeyStr].serie || [];
    for (const s of serie) {
      const [dd, mm, yyyy] = s.data.split('/');
      const dt = new Date(+yyyy, +mm - 1, +dd);
      if (dt >= inicio && dt <= fim) { total += (s.hospedes || 0); temDados = true; }
    }
  }
  return temDados ? total : null;
}

app.get('/api/concessionarias', async (req, res) => {
  const tipo = (req.query.tipo || 'gasmig').toLowerCase();
  try {
    if (tipo !== 'gasmig') return res.json({ tipo, periodos: [], aviso: 'Ainda não configurado.' });

    const arquivos = await listaArquivosConsumo('GASMIG');
    if (!arquivos.length) return res.json({ tipo, periodos: [], aviso: 'Nenhum PDF da GASMIG encontrado na pasta.' });

    const fmt = d => `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;

    // Baixa TODAS as faturas — cada uma traz o próprio período atual + até 12 meses de histórico.
    // Junta tudo (sem duplicar), já que faturas mais antigas alcançam meses que a mais recente não cobre.
    const valoresPorPeriodo = {};
    const periodosPorLabel = {};
    for (const f of arquivos) {
      try {
        const bufF = await downloadFile(f.id);
        const textoF = await pdfParse(bufF).then(r => r.text);
        const periodosArquivo = parseGasmigTexto(textoF);
        const pF = periodosArquivo[0]; // período atual dessa fatura específica
        if (pF && pF.totalPagar !== null) valoresPorPeriodo[`${fmt(pF.inicio)} a ${fmt(pF.fim)}`] = pF.totalPagar;
        for (const p of periodosArquivo) {
          const label = `${fmt(p.inicio)} a ${fmt(p.fim)}`;
          if (!periodosPorLabel[label]) periodosPorLabel[label] = p;
        }
      } catch(e) { console.warn('[GASMIG]', f.name, e.message); }
    }

    const periodosBrutos = Object.values(periodosPorLabel).sort((a, b) => b.inicio - a.inicio);

    const store = fs.existsSync(RESULT_FILE) ? JSON.parse(fs.readFileSync(RESULT_FILE, 'utf8')) : { dados: {} };

    const periodos = periodosBrutos.map(p => {
      const ocupacao = somaOcupacaoPeriodo(store, p.inicio, p.fim);
      const periodoLabel = `${fmt(p.inicio)} a ${fmt(p.fim)}`;
      const totalPagar = valoresPorPeriodo[periodoLabel] ?? null;
      return {
        periodo: periodoLabel,
        consumo: +p.consumo.toFixed(3),
        ocupacao,
        consumoPerCapita: ocupacao > 0 ? +(p.consumo / ocupacao).toFixed(3) : null,
        valorPerCapita: (ocupacao > 0 && totalPagar !== null) ? +(totalPagar / ocupacao).toFixed(2) : null
      };
    });

    res.json({ tipo, arquivo: arquivos.map(a => a.name).join(', '), periodos });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Converte serial Excel (número) em dd/mm/yyyy; se já vier como texto, devolve como está
function formataDataExcel(v) {
  if (typeof v === 'number') {
    const dt = new Date((v - 25569) * 86400 * 1000);
    return `${String(dt.getUTCDate()).padStart(2,'0')}/${String(dt.getUTCMonth()+1).padStart(2,'0')}/${dt.getUTCFullYear()}`;
  }
  return String(v || '');
}

// Carrega o arquivo INSUMOS_organizado.xlsx e devolve só as linhas de Saída (únicas que têm
// Centro de Custo real — nas linhas de Recebimento essa coluna traz o nome do fornecedor)
async function carregarItensCustos() {
  const buf = await downloadFile(CUSTOS_FILE_ID);
  const wb = XLSX.read(buf, { type: 'buffer' });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
  const itens = [];
  let idx = 0;
  for (const r of rows) {
    if (String(r['E/S/H']).trim() !== 'Saída') continue;
    itens.push({
      id: `s${idx}`,
      dataMovimento: formataDataExcel(r['Data Movimento']),
      artigo: String(r['Artigo'] || ''),
      produto: String(r['Produto'] || ''),
      documento: String(r['Documento'] || ''),
      saidaQtd: Number(r['Saída (Qtd)']) || 0,
      saidaValor: +(Number(r['Saída (Valor)']) || 0).toFixed(2),
      valorMedio: +(Number(r['Valor (Preço Médio)']) || 0).toFixed(2),
      centroCustoOriginal: String(r['Centro de Custo'] || '').trim(),
      dataLanc: formataDataExcel(r['Data Lanc.'])
    });
    idx++;
  }
  return itens;
}

function aplicarCorrecoes(itens, correcoes) {
  return itens
    .filter(it => !correcoes.auditoria[it.id]) // remove os que estão em auditoria
    .map(it => ({
      ...it,
      centroCusto: correcoes.reatribuidos[it.id] || it.centroCustoOriginal
    }));
}

app.get('/api/custos/centros', async (req, res) => {
  try {
    const itens = aplicarCorrecoes(await carregarItensCustos(), loadCorrecoes());
    const porCentro = {};
    for (const it of itens) porCentro[it.centroCusto] = (porCentro[it.centroCusto] || 0) + 1;
    const centros = Object.entries(porCentro).map(([nome, qtd]) => ({ nome, qtd })).sort((a,b) => b.qtd - a.qtd);
    res.json({ centros });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/custos/itens', async (req, res) => {
  try {
    const centro = req.query.centro;
    if (!centro) return res.status(400).json({ error: 'Parâmetro centro é obrigatório.' });
    const itens = aplicarCorrecoes(await carregarItensCustos(), loadCorrecoes());
    res.json({ itens: itens.filter(it => it.centroCusto === centro) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/custos/auditoria', async (req, res) => {
  try {
    const correcoes = loadCorrecoes();
    const todos = await carregarItensCustos();
    const emAuditoria = todos
      .filter(it => correcoes.auditoria[it.id])
      .map(it => ({ ...it, centroCustoOriginal: correcoes.auditoria[it.id].centroOriginal || it.centroCustoOriginal }));
    res.json({ itens: emAuditoria });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/custos/mover-auditoria', async (req, res) => {
  try {
    const { id, centroAtual } = req.body;
    if (!id) return res.status(400).json({ error: 'id é obrigatório.' });
    const correcoes = loadCorrecoes();
    correcoes.auditoria[id] = { centroOriginal: centroAtual || '', movidoEm: new Date().toISOString() };
    delete correcoes.reatribuidos[id];
    saveCorrecoes(correcoes);
    res.json({ ok: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/custos/corrigir', async (req, res) => {
  try {
    const { id, novoCentro } = req.body;
    if (!id || !novoCentro) return res.status(400).json({ error: 'id e novoCentro são obrigatórios.' });
    const correcoes = loadCorrecoes();
    delete correcoes.auditoria[id];
    correcoes.reatribuidos[id] = novoCentro;
    saveCorrecoes(correcoes);
    res.json({ ok: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

const MES_ABREV_NUM = { jan:1, fev:2, mar:3, abr:4, mai:5, jun:6, jul:7, ago:8, set:9, out:10, nov:11, dez:12 };
const MES_NOMES = ['','Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

function parseTripadvisorTexto(texto, ano) {
  const linhas = texto.split('\n');
  const dados = [];
  for (const linha of linhas) {
    const m = linha.trim().match(/^([a-zà-ú]{3})\s*-\s*(\d+)\s*-\s*([\d,]+)$/i);
    if (!m) continue;
    const abrev = m[1].toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'');
    const mesNum = MES_ABREV_NUM[abrev];
    if (!mesNum) continue;
    dados.push({
      mes: `${ano}-${String(mesNum).padStart(2,'0')}`,
      mesLabel: MES_NOMES[mesNum],
      avaliacoes: parseInt(m[2], 10),
      nota: parseFloat(m[3].replace(',', '.'))
    });
  }
  dados.sort((a, b) => a.mes.localeCompare(b.mes));
  return dados;
}

app.get('/api/tripadvisor', async (req, res) => {
  try {
    const ano = req.query.ano || new Date().getFullYear();
    const texto = await downloadGoogleDocText(TRIPADVISOR_DOC_ID);
    const dados = parseTripadvisorTexto(texto, ano);
    res.json({ ano, dados });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/debug-tripadvisor', async (req, res) => {
  try {
    const texto = await downloadGoogleDocText(TRIPADVISOR_DOC_ID);
    res.json({ tamanho: texto.length, texto });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/debug-custos-buscar', async (req, res) => {
  try {
    const termo = normalizaTexto(req.query.q || 'file');
    const buf = await downloadFile(CUSTOS_FILE_ID);
    const wb = XLSX.read(buf, { type: 'buffer' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    const encontrados = rows.filter(r => normalizaTexto(r['Produto']).includes(termo));
    res.json({ termo, totalEncontrado: encontrados.length, linhas: encontrados.slice(0, 30) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/debug-custos', async (req, res) => {
  try {
    const buf = await downloadFile(CUSTOS_FILE_ID);
    const wb = XLSX.read(buf, { type: 'buffer' });
    const sheetName = wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
    res.json({
      aba: sheetName,
      totalLinhas: rows.length,
      colunas: rows.length ? Object.keys(rows[0]) : [],
      primeirasLinhas: rows.slice(0, 5)
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/debug-buscar-arquivo', async (req, res) => {
  try {
    const nome = req.query.nome || 'INSUMOS_organizado';
    const q = encodeURIComponent(`name contains '${nome}' and trashed=false`);
    const d = await driveGet(`files?q=${q}&fields=files(id,name,mimeType,modifiedTime,parents)&corpora=allDrives`);
    res.json({ encontrados: d.files || [] });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/debug-consumo', async (req, res) => {
  try {
    const q = encodeURIComponent(`'${CONSUMO_FOLDER_ID}' in parents and trashed=false and mimeType='application/pdf'`);
    const d = await driveGet(`files?q=${q}&fields=files(id,name,mimeType,modifiedTime)&corpora=allDrives`);
    const arquivos = (d.files || []).map(f => ({ ...f, mesKey: parseConsumoFilename(f.name) })).filter(f => f.mesKey);
    if (!arquivos.length) return res.json({ arquivos: [], aviso: 'Nenhum PDF com padrão de nome MMAAAA encontrado.' });

    arquivos.sort((a,b) => b.mesKey.localeCompare(a.mesKey));
    const maisRecente = arquivos[0];
    const buf = await downloadFile(maisRecente.id);
    const parsed = await pdfParse(buf);

    res.json({ arquivos, maisRecente: maisRecente.name, mesKey: maisRecente.mesKey, texto: parsed.text });
  } catch(e) { res.status(500).json({ erro: e.message }); }
});

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

app.get('/api/debug-produtos', async (req, res) => {
  try {
    const vendaDir = await findFolder(SHARED_DRIVE, 'VENDAS');
    const vendaPdf = await latestPdf(vendaDir.id);
    const vendaBuf = await downloadFile(vendaPdf.id);
    const vendaText = await pdfParse(vendaBuf).then(r => r.text);
    const linhas = vendaText.split('\n');
    const kw = /CERVEJA|AGUA|ÁGUA|REFRIGERANTE|SUCO|VINHO|DOSE|DRINK|WHISK|VODKA|GIN|CHOPP|ENERG|REFRI|COCA|GUARANA|ESPUMANTE|CAIPIR/i;
    const amostra = linhas.filter(l => kw.test(l));
    res.json({
      arquivo: vendaPdf.name,
      totalLinhas: linhas.length,
      primeiras60: linhas.slice(0, 60),
      amostraBebidas: amostra.slice(0, 60)
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Relatório de vinhos (garrafas e taças) desde uma data ────────────────────
// Palavras "significativas" do nome (ignora conectivos/descrições genéricas) — usadas
// para casar o nome do vinho na base de estoque com o nome extraído do PDV
const STOPWORDS_VINHO = new Set(['DE','DO','DA','DOS','DAS','COM','TINTO','BRANCO','ROSE','ROSADO','UN','GF','TAÇA','TACA','VINHO','VINHOS','VIN','750ML','375ML','ESPUMANTE','ESPUM','ML','LT']);
function palavrasSignificativas(nome) {
  return normalizaTexto(nome).split(/\s+/).filter(p => p.length >= 3 && !STOPWORDS_VINHO.has(p));
}

// Apura as vendas de vinho do mês (a partir do texto do PDF de vendas) e atualiza o campo
// "vendas" (negativo) de cada item da base de estoque. Casa cada linha do PDV com o item da
// base que tiver MAIS palavras em comum (não apenas a primeira), para não duplicar a mesma
// venda em vários itens que compartilham a marca (ex: "Arg Cordero..." em 3 variações)
function apurarVendasVinhosDoMes(texto) {
  const itensPdv = extrairItensVinho(texto); // [{data, nome, tipo, qtd}]
  const itensEstoque = loadVinhos();
  const palavrasPorItem = itensEstoque.map(it => palavrasSignificativas(it.nome));

  const porCodigo = {};
  let tacaTinto = 0, tacaBranco = 0;
  for (const it of itensPdv) {
    if (it.tipo === 'Taça') {
      if (/TINTO/.test(it.nome)) { tacaTinto += it.qtd; continue; }
      if (/BRANCO/.test(it.nome)) { tacaBranco += it.qtd; continue; }
    }
    const palavrasPdv = new Set(palavrasSignificativas(it.nome));
    if (!palavrasPdv.size) continue;

    let melhorIdx = -1, melhorScore = 0;
    itensEstoque.forEach((_, idx) => {
      const score = palavrasPorItem[idx].reduce((s, p) => s + (palavrasPdv.has(p) ? 1 : 0), 0);
      if (score > melhorScore) { melhorScore = score; melhorIdx = idx; }
    });
    if (melhorIdx === -1) continue; // nenhuma correspondência — não atribui a ninguém

    const codigo = itensEstoque[melhorIdx].codigo;
    porCodigo[codigo] = (porCodigo[codigo] || 0) + it.qtd;
  }

  for (const item of itensEstoque) {
    const qtdVendida = porCodigo[item.codigo] || 0;
    item.vendas = qtdVendida > 0 ? -Math.round(qtdVendida) : 0;
  }
  saveVinhos(itensEstoque);
  fs.writeFileSync(BEBIDAS_TACA_FILE, JSON.stringify({
    tacaTinto: Math.round(tacaTinto), tacaBranco: Math.round(tacaBranco), atualizadoEm: new Date().toISOString()
  }, null, 2));
}

function extrairItensVinho(texto) {
  const linhas = texto.split('\n');
  const itens = [];
  let lastDate = null;
  let grupoAtual = null;

  for (let i = 0; i < linhas.length; i++) {
    const line = linhas[i].trim();

    const gm = line.match(/^GRUPO:\s*(.+)$/i);
    if (gm) { grupoAtual = gm[1].trim().toUpperCase(); continue; }

    const dm = line.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (dm) { lastDate = new Date(+dm[3], +dm[2] - 1, +dm[1]); continue; }

    const qtdR = (line.match(/R\$/g) || []).length;
    if (grupoAtual === 'VINHOS' && line.startsWith('R$') && qtdR >= 4) {
      // Junta as linhas anteriores (até 3) para formar o nome do item, parando em limites conhecidos
      let nomeParts = [];
      let j = i - 1;
      while (j >= 0 && nomeParts.length < 3) {
        const prev = linhas[j].trim();
        if (!prev) { j--; continue; }
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(prev)) break;
        if (/^\d{2}:\d{2}$/.test(prev)) break;
        if (/Mesa|Posi[çc][ãa]o/i.test(prev)) break;
        if (prev.startsWith('R$')) break;
        // Fragmento curto isolado (ex: "AL") — continuação de "Posição N -" quebrada em linha própria, não é nome
        if (/^[A-ZÀ-Ú]{1,3}$/i.test(prev)) { j--; continue; }
        nomeParts.unshift(prev);
        j--;
      }
      let nome = nomeParts.join(' ').replace(/\s+/g, ' ').trim().toUpperCase();
      // Corrige truncamento de origem: a impressão do PDV corta "VINHO BRANCO" em "VINHO BRANC"
      nome = nome.replace(/VINHO BRANC$/, 'VINHO BRANCO');
      // Descarta nomes sem nenhuma letra (fragmento numérico/monetário capturado por engano,
      // geralmente por quebra de página que atropela a estrutura normal da linha)
      const temLetra = /[A-ZÀ-Ú]/.test(nome);
      if (nome && temLetra) {
        // Preço e QTD vêm concatenados sem separador (ex: "R$ 29,001,00"); QTD é sempre o último
        // número válido "d{2}" nesse trecho — o preço nunca ancora corretamente por causa da ambiguidade
        const primeiroSeg = line.split('R$')[1] || '';
        const matches = [...primeiroSeg.matchAll(/\d{1,3}(?:\.\d{3})*,\d{2}(?!\d)/g)];
        const qtd = matches.length ? parseFloat(matches[matches.length - 1][0].replace(/\./g,'').replace(',','.')) : null;
        // Sanity check: nenhuma venda unitária de vinho passa de algumas dezenas de unidades
        if (qtd !== null && qtd > 0 && qtd <= 50 && lastDate) {
          // Tudo em GRUPO:VINHOS que não é vendido em taça é garrafa (vendida pelo nome/marca do vinho)
          const tipo = /TA[ÇC]A/.test(nome) ? 'Taça' : 'Garrafa';
          itens.push({ data: lastDate, nome, tipo, qtd });
        }
      }
    }
  }
  return itens;
}

app.get('/api/debug-apurar-vinhos', async (req, res) => {
  try {
    const vendaDir = await findFolder(SHARED_DRIVE, 'VENDAS');
    if (!vendaDir) return res.status(404).json({ error: 'Pasta VENDAS não encontrada.' });
    const pdfs = await allPdfs(vendaDir.id);
    const mesAtual = mesAtualKey();
    const arquivosDoMes = pdfs.filter(f => mesKey(f.name) === mesAtual);
    if (!arquivosDoMes.length) return res.json({ mesAtual, erro: 'Nenhum PDF de vendas encontrado para o mês corrente.' });
    // usa o mesmo critério do sincronizar(): o mais recentemente modificado
    const arquivo = arquivosDoMes.sort((a,b) => (a.modifiedTime > b.modifiedTime ? -1 : 1))[0];

    const buf = await downloadFile(arquivo.id);
    const texto = await pdfParse(buf).then(r => r.text);

    const antes = loadVinhos().map(it => ({ codigo: it.codigo, nome: it.nome, vendas: it.vendas }));
    apurarVendasVinhosDoMes(texto);
    const depois = loadVinhos().map(it => ({ codigo: it.codigo, nome: it.nome, vendas: it.vendas }));
    const tacas = loadTacas();

    res.json({ arquivo: arquivo.name, mesAtual, ok: true, tacas, antes, depois });
  } catch(e) {
    res.status(500).json({ ok: false, erro: e.message, stack: e.stack });
  }
});

app.get('/api/debug-vinhos', async (req, res) => {
  try {
    const vendaDir = await findFolder(SHARED_DRIVE, 'VENDAS');
    const arquivoNome = req.query.arquivo;
    const pdfs = await allPdfs(vendaDir.id);
    const pdf = arquivoNome ? pdfs.find(f => f.name === arquivoNome) : pdfs.find(f => mesKey(f.name) === '2026-07');
    if (!pdf) return res.status(404).json({ error: 'Arquivo não encontrado.', disponiveis: pdfs.map(f=>f.name) });
    const buf = await downloadFile(pdf.id);
    const texto = await pdfParse(buf).then(r => r.text);
    const linhas = texto.split('\n');

    if (req.query.anomalias) {
      const esperado = /^TA[ÇC]A DE VINHO (TINTO|BRANCO)$/;
      const anomalias = [];
      for (let i = 0; i < linhas.length; i++) {
        const line = linhas[i].trim();
        const qtdR = (line.match(/R\$/g) || []).length;
        if (line.startsWith('R$') && qtdR >= 4) {
          let nomeParts = [];
          let j = i - 1;
          while (j >= 0 && nomeParts.length < 3) {
            const prev = linhas[j].trim();
            if (!prev) { j--; continue; }
            if (/^\d{2}\/\d{2}\/\d{4}$/.test(prev)) break;
            if (/^\d{2}:\d{2}$/.test(prev)) break;
            if (/Mesa|Posi[çc][ãa]o/i.test(prev)) break;
            if (prev.startsWith('R$')) break;
            nomeParts.unshift(prev);
            j--;
          }
          const nome = nomeParts.join(' ').replace(/\s+/g, ' ').trim().toUpperCase();
          if (/VINHO/.test(nome) && !esperado.test(nome)) {
            anomalias.push({ idx: i, nome, contexto: linhas.slice(Math.max(0,i-8), i+3).map((l,k)=>`${Math.max(0,i-8)+k}: ${l}`) });
          }
        }
      }
      return res.json({ arquivo: pdf.name, totalAnomalias: anomalias.length, anomalias });
    }

    const ocorrencias = [];
    for (let i = 0; i < linhas.length; i++) {
      if (/VINHO/i.test(linhas[i])) {
        ocorrencias.push({ idx: i, contexto: linhas.slice(Math.max(0,i-6), i+6).map((l,k)=>`${Math.max(0,i-6)+k}: ${l}`) });
      }
    }
    res.json({ arquivo: pdf.name, totalOcorrencias: ocorrencias.length, ocorrencias: ocorrencias.slice(0, 15) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/relatorio-vinhos', async (req, res) => {
  try {
    const mesesFiltro = req.query.meses ? req.query.meses.split(',').map(s => s.trim()) : null;
    const desde = req.query.desde ? new Date(req.query.desde) : new Date(2026, 3, 1); // 1º de abril de 2026 por padrão

    const vendaDir = await findFolder(SHARED_DRIVE, 'VENDAS');
    if (!vendaDir) return res.status(404).json({ error: 'Pasta VENDAS não encontrada.' });
    const pdfs = await allPdfs(vendaDir.id);
    const pdfsRelevantes = pdfs.filter(f => {
      const k = mesKey(f.name);
      if (!k) return false;
      if (mesesFiltro) return mesesFiltro.includes(k);
      return k >= `${desde.getFullYear()}-${String(desde.getMonth()+1).padStart(2,'0')}`;
    });

    const todosItens = [];
    const arquivosProcessados = [];
    for (const f of pdfsRelevantes) {
      const buf = await downloadFile(f.id);
      const texto = await pdfParse(buf).then(r => r.text);
      const itens = extrairItensVinho(texto).filter(it => mesesFiltro ? true : it.data >= desde);
      todosItens.push(...itens);
      arquivosProcessados.push({ arquivo: f.name, itensEncontrados: itens.length });
    }

    // Agrega por item + mês
    const porItem = {};
    const porMes = {};
    for (const it of todosItens) {
      if (!porItem[it.nome]) porItem[it.nome] = { nome: it.nome, tipo: it.tipo, qtdTotal: 0, porMes: {} };
      porItem[it.nome].qtdTotal += it.qtd;
      const mesKeyStr = `${it.data.getFullYear()}-${String(it.data.getMonth()+1).padStart(2,'0')}`;
      porItem[it.nome].porMes[mesKeyStr] = (porItem[it.nome].porMes[mesKeyStr] || 0) + it.qtd;
      if (!porMes[mesKeyStr]) porMes[mesKeyStr] = { garrafas: 0, tacas: 0 };
      if (it.tipo === 'Garrafa') porMes[mesKeyStr].garrafas += it.qtd;
      else if (it.tipo === 'Taça') porMes[mesKeyStr].tacas += it.qtd;
    }

    res.json({
      desde: desde.toISOString().slice(0,10),
      arquivosProcessados,
      itens: Object.values(porItem).sort((a,b) => b.qtdTotal - a.qtdTotal),
      resumoMensal: porMes
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

app.get('/eventos', (req, res) => res.sendFile(path.join(__dirname, 'public', 'eventos.html')));
app.get('/caed', (req, res) => res.redirect('/licitacoes'));
app.get('/licitacoes', (req, res) => res.sendFile(path.join(__dirname, 'public', 'licitacoes.html')));
app.get('/concessionarias', (req, res) => res.sendFile(path.join(__dirname, 'public', 'concessionarias.html')));
app.get('/custos', (req, res) => res.sendFile(path.join(__dirname, 'public', 'custos.html')));
app.get('/tripadvisor', (req, res) => res.sendFile(path.join(__dirname, 'public', 'tripadvisor.html')));
app.get('/bebidas', (req, res) => res.sendFile(path.join(__dirname, 'public', 'bebidas.html')));

function comEstoqueFinal(it) {
  // vendas é negativo (saída), então soma normalmente
  return { ...it, estoqueFinal: +(it.estoqueInicial + (it.entradas||0) + (it.vendas||0)).toFixed(3) };
}

function loadTacas() {
  try { return JSON.parse(fs.readFileSync(BEBIDAS_TACA_FILE, 'utf8')); }
  catch { return { tacaTinto: 0, tacaBranco: 0, atualizadoEm: null }; }
}

app.get('/api/bebidas/vinhos', (req, res) => {
  const itens = loadVinhos().map(comEstoqueFinal).sort((a,b) => a.nome.localeCompare(b.nome, 'pt-BR'));
  const { tacaTinto, tacaBranco } = loadTacas();
  res.json({ itens, tacaTinto, tacaBranco });
});

app.post('/api/bebidas/vinhos/novo', (req, res) => {
  const { codigo, nome, tamanho, categoria, estoqueInicial } = req.body;
  const codigoLimpo = String(codigo || '').trim();
  const nomeLimpo = String(nome || '').trim();
  if (!codigoLimpo || !nomeLimpo) return res.status(400).json({ error: 'Código e nome são obrigatórios.' });

  const itens = loadVinhos();
  if (itens.some(i => i.codigo === codigoLimpo)) {
    return res.status(409).json({ error: `Já existe um vinho cadastrado com o código ${codigoLimpo}.` });
  }

  const novo = {
    codigo: codigoLimpo,
    nome: nomeLimpo.toLowerCase(),
    tamanho: String(tamanho || 'UN').trim(),
    categoria: String(categoria || ''),
    estoqueInicial: +estoqueInicial || 0,
    vendas: 0,
    entradas: 0,
    auditoria: null,
    observacao: ''
  };
  itens.push(novo);
  saveVinhos(itens);
  res.json({ ok: true, item: comEstoqueFinal(novo) });
});

app.get('/api/bebidas/vinhos/auditoria', (req, res) => {
  const itens = loadVinhos().map(comEstoqueFinal)
    .filter(it => it.auditoria !== null && it.auditoria !== undefined && +it.auditoria !== it.estoqueFinal)
    .map(it => ({ ...it, diferenca: +((+it.auditoria) - it.estoqueFinal).toFixed(3) }))
    .sort((a,b) => a.nome.localeCompare(b.nome, 'pt-BR'));
  res.json({ itens });
});

const CAMPOS_TEXTO_VINHO = ['observacao', 'nome', 'tamanho', 'categoria'];
app.post('/api/bebidas/vinhos/atualizar', (req, res) => {
  const { codigo, campo, valor } = req.body;
  if (!codigo || !['vendas','entradas','auditoria', ...CAMPOS_TEXTO_VINHO].includes(campo)) {
    return res.status(400).json({ error: 'codigo e campo (vendas|entradas|auditoria|observacao|nome|tamanho|categoria) são obrigatórios.' });
  }
  const itens = loadVinhos();
  const item = itens.find(i => i.codigo === codigo);
  if (!item) return res.status(404).json({ error: 'Item não encontrado.' });
  if (CAMPOS_TEXTO_VINHO.includes(campo)) {
    item[campo] = String(valor || '');
  } else {
    item[campo] = valor === '' || valor === null ? (campo === 'auditoria' ? null : 0) : +valor;
  }
  saveVinhos(itens);
  res.json({ ok: true, item: comEstoqueFinal(item) });
});

function mesAtualKey() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

app.get('/api/eventos', (req, res) => {
  if (!fs.existsSync(RESULT_FILE)) return res.status(404).json({ error: 'Sem dados. Clique em Sincronizar.' });
  const store = JSON.parse(fs.readFileSync(RESULT_FILE, 'utf8'));
  const meses = [...new Set([...(store.meses || []), ...Object.keys(store.eventosRaw || {})])].sort().reverse();
  const mesAtual = mesAtualKey();
  const mes = req.query.mes || (meses.includes(mesAtual) ? mesAtual : meses[0]);
  if (!mes) return res.json({ meses: [], mes: null, linhas: [], totais: {} });
  const raw = store.eventosRaw?.[mes];
  if (!raw) return res.json({ meses, mes, linhas: [], totais: { pax:0, banq:0, sala:0, equip:0, total:0 } });

  const totais = { pax: raw.pax, banq: raw.banq, sala: raw.sala, equip: raw.equip, total: raw.total };

  // Marca cada evento como já realizado ou ainda agendado (só é relevante no mês corrente)
  const hoje = new Date(); hoje.setHours(23, 59, 59, 999);
  const linhas = (raw.linhas || []).map(l => {
    let status = 'realizado';
    if (l.data) {
      const [dd, mm, yyyy] = l.data.split('/');
      if (new Date(+yyyy, +mm - 1, +dd) > hoje) status = 'agendado';
    }
    return { ...l, status };
  });

  const realizadoAteHoje = mes === mesAtual
    ? linhas.filter(l => l.status === 'realizado').reduce((acc, l) => ({
        pax:   acc.pax   + (l.pax   || 0),
        banq:  acc.banq  + (l.banq  || 0),
        sala:  acc.sala  + (l.sala  || 0),
        equip: acc.equip + (l.equip || 0),
        total: acc.total + (l.total || 0)
      }), { pax: 0, banq: 0, sala: 0, equip: 0, total: 0 })
    : null;

  res.json({ meses, mes, mesLabel: mesLabel(mes), linhas, totais, realizadoAteHoje });
});

app.get('/api/caed', (req, res) => {
  if (!fs.existsSync(RESULT_FILE)) return res.status(404).json({ error: 'Sem dados. Clique em Sincronizar.' });
  const store = JSON.parse(fs.readFileSync(RESULT_FILE, 'utf8'));
  const meses = [...new Set([...(store.meses || []), ...Object.keys(store.caedRaw || {})])].sort().reverse();
  const mesAtual = mesAtualKey();
  const mes = req.query.mes || (meses.includes(mesAtual) ? mesAtual : meses[0]);
  if (!mes) return res.json({ meses: [], mes: null, linhas: [], totais: {} });
  const raw = store.caedRaw?.[mes];
  if (!raw) return res.json({ meses, mes, linhas: [], totais: { pessoas:0, semPensao:0, meiaPensao:0, pensaoCompleta:0, faturamento:0 } });

  const totais = {
    pessoas: raw.totalPessoas, semPensao: raw.semPensao, meiaPensao: raw.meiaPensao,
    pensaoCompleta: raw.pensaoCompleta, faturamento: raw.totalFaturamento
  };

  // Marca cada hóspede como agendado (check-in futuro), em andamento (já entrou, ainda não saiu)
  // ou realizado (check-out já passou) — só relevante para o mês corrente
  const hoje = new Date(); hoje.setHours(23, 59, 59, 999);
  const linhas = (raw.linhas || []).map(l => {
    let status = 'realizado';
    if (l.checkin) {
      const [ddIn, mmIn, yyyyIn] = l.checkin.split('/');
      const dataCheckin = new Date(+yyyyIn, +mmIn - 1, +ddIn);
      if (dataCheckin > hoje) {
        status = 'agendado';
      } else if (l.checkout) {
        const [ddOut, mmOut, yyyyOut] = l.checkout.split('/');
        const dataCheckout = new Date(+yyyyOut, +mmOut - 1, +ddOut);
        if (dataCheckout > hoje) status = 'em_andamento';
      }
    }
    return { ...l, status };
  });

  const realizadoAteHoje = mes === mesAtual
    ? (() => {
        let pessoas = 0, faturamento = 0;
        for (const [dataStr, v] of Object.entries(raw.daily || {})) {
          const [dd, mm, yyyy] = dataStr.split('/');
          if (new Date(+yyyy, +mm - 1, +dd) <= hoje) { pessoas += v.pessoas; faturamento += v.faturamento; }
        }
        return { pessoas, faturamento: +faturamento.toFixed(2) };
      })()
    : null;

  const daily = Object.entries(raw.daily || {})
    .map(([data, v]) => ({ data, semPensao: v.semPensao, meiaPensao: v.meiaPensao, pensaoCompleta: v.pensaoCompleta }))
    .sort((a, b) => a.data.split('/').reverse().join('').localeCompare(b.data.split('/').reverse().join('')));

  res.json({ meses, mes, mesLabel: mesLabel(mes), linhas, totais, realizadoAteHoje, daily });
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
