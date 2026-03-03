/**
 * CRM Corretagem - Code.gs (COMPLETO E ATUALIZADO)
 * - Web App com roteamento por ?page=
 * - Includes compartilhados (CSS/JS)
 * - Ajuste de schema: garante coluna "id" em Agenda_Visitas
 * - Menus opcionais no Sheets
 * - Wrappers globais para google.script.run (DataService + VisitService + Dashboard)
 *
 * IMPORTANTÍSSIMO:
 * - Este arquivo NÃO pode conter HTML.
 * - HTML deve ficar somente nos arquivos .html (ex.: Form_BaseClientes.html).
 */

const CFG = {
  TITLE: "CRM Corretagem",
  DEFAULT_PAGE: "Form_Menu",
  PAGES: {
    "Form_Menu": true,
    "Form_BaseClientes": true,
    "Form_Estoque": true,
    "Form_LeadsCompradores": true,
    "Form_LeadsVendedores": true,
    "Form_AgendarVisita": true,
    "Form_RegistrarVisita": true,
    "Form_PDFVisitas": true
  }
};

/**
 * Leads_Compradores - helper de parsing DD/MM/AAAA (e compatíveis)
 * Retorna Date (00:00) ou null se inválido.
 */
function parseDDMMYYYY_(v) {
  if (v === null || v === undefined || v === "") return null;

  // Se já é Date
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  const s = String(v).trim();
  if (!s) return null;

  // Aceita DD/MM/AAAA
  let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    const dd = Number(m[1]), mm = Number(m[2]), yyyy = Number(m[3]);
    if (yyyy < 1900 || yyyy > 2100) return null;
    if (mm < 1 || mm > 12) return null;
    const d = new Date(yyyy, mm - 1, dd);
    // valida (evita 31/02 etc)
    if (d.getFullYear() !== yyyy || d.getMonth() !== (mm - 1) || d.getDate() !== dd) return null;
    return d;
  }

  // Aceita YYYY-MM-DD
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const yyyy = Number(m[1]), mm = Number(m[2]), dd = Number(m[3]);
    if (yyyy < 1900 || yyyy > 2100) return null;
    if (mm < 1 || mm > 12) return null;
    const d = new Date(yyyy, mm - 1, dd);
    if (d.getFullYear() !== yyyy || d.getMonth() !== (mm - 1) || d.getDate() !== dd) return null;
    return d;
  }

  return null;
}

/**
 * (Opcional, mas ajuda) Normaliza início/fim do dia para filtro por intervalo
 */
function dayStart_(d) {
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}
function dayEnd_(d) {
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
}

/**
 * parseAnyDate_
 * Converte vários formatos em Date (normalizado 00:00:00)
 * Aceita:
 *   - DD/MM/AAAA
 *   - AAAA-MM-DD
 *   - Date object (Sheets)
 *   - string parseável
 * Retorna Date ou null
 */
function parseAnyDate_(v) {
  if (v === null || v === undefined || v === "") return null;

  // 1️⃣ Se já for Date válido
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  const s = String(v).trim();
  if (!s) return null;

  // 2️⃣ DD/MM/AAAA
  let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    const dd = Number(m[1]);
    const mm = Number(m[2]);
    const yyyy = Number(m[3]);

    if (yyyy < 1900 || yyyy > 2100) return null;
    if (mm < 1 || mm > 12) return null;

    const d = new Date(yyyy, mm - 1, dd);

    if (
      d.getFullYear() !== yyyy ||
      d.getMonth() !== (mm - 1) ||
      d.getDate() !== dd
    ) return null;

    return d;
  }

  // 3️⃣ AAAA-MM-DD
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const yyyy = Number(m[1]);
    const mm = Number(m[2]);
    const dd = Number(m[3]);

    if (yyyy < 1900 || yyyy > 2100) return null;
    if (mm < 1 || mm > 12) return null;

    const d = new Date(yyyy, mm - 1, dd);

    if (
      d.getFullYear() !== yyyy ||
      d.getMonth() !== (mm - 1) ||
      d.getDate() !== dd
    ) return null;

    return d;
  }

  // 4️⃣ Fallback (última tentativa)
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  return null;
}

function doGet(e) {
  ensureSchema_();

  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : CFG.DEFAULT_PAGE;
  const safePage = CFG.PAGES[page] ? page : CFG.DEFAULT_PAGE;

  return HtmlService
    .createTemplateFromFile(safePage)
    .evaluate()
    .setTitle(CFG.TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function onOpen() {
  try {
    ensureSchema_();
    SpreadsheetApp.getUi()
      .createMenu("CRM Corretagem")
      .addItem("Abrir Web App", "openWebApp_")
      .addSeparator()
      .addItem("Rodar ajuste de estrutura", "ensureSchema_")
      .addToUi();
  } catch (err) {
    // silencioso para não travar
  }
}

function openWebApp_() {
  const url = getAppUrl();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial">
      <p><b>Web App:</b></p>
      <p><a href="${url}" target="_blank">${url}</a></p>
    </div>
  `).setWidth(520).setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, "Abrir CRM Corretagem");
}

/**
 * ✅ Garante que a aba "Agenda_Visitas" tenha coluna "id" (minúsculo) na coluna A.
 * - Se existir "id" OU "ID", não altera.
 * - Se não existir, cria e preenche ids sequenciais.
 */
function ensureSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Agenda_Visitas");
  if (!sh) return;

  const lc = sh.getLastColumn();
  if (lc < 1) return;

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
  const hasIdLower = headers.includes("id");
  const hasIdUpper = headers.includes("ID");
  if (hasIdLower || hasIdUpper) return;

  sh.insertColumnBefore(1);
  sh.getRange(1, 1).setValue("id");
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight("bold");

  const lr = sh.getLastRow();
  if (lr >= 2) {
    const idRange = sh.getRange(2, 1, lr - 1, 1);
    const ids = idRange.getValues();
    let next = 1;

    for (let i = 0; i < ids.length; i++) {
      if (!ids[i][0]) ids[i][0] = next++;
      else {
        const n = Number(ids[i][0]);
        if (!isNaN(n) && isFinite(n)) next = Math.max(next, n + 1);
      }
    }
    idRange.setValues(ids);
  }
}

/* ===============================
   DataService wrappers
================================ */

function getMeta(sheetName) { return DataService.getMeta(sheetName); }
function listRecords(sheetName, idCol, labelCols) { return DataService.listRecords(sheetName, idCol, labelCols); }
function getById(sheetName, idCol, idVal) { return DataService.getById(sheetName, idCol, idVal); }
function upsertById(sheetName, idCol, obj) { return DataService.upsertById(sheetName, idCol, obj); }
function deleteById(sheetName, idCol, idVal) { return DataService.deleteById(sheetName, idCol, idVal); }
function getNextNumericId(sheetName, idCol) { return DataService.getNextNumericId(sheetName, idCol); }
function listClientesForSelect() { return DataService.listClientesForSelect(); }

/* ===============================
   VisitService wrappers (prefix VS_)
================================ */

function VS_listAgendaForSelect() { return listAgendaForSelect(); }
function VS_getAgendaById(idAgenda) { return getAgendaById(idAgenda); } // upgrade natural
function VS_createAgendaVisit(obj) { return createAgendaVisit(obj); }
function VS_saveFatoVisita(obj) { return saveFatoVisita(obj); }
function VS_getFatoVisitaByIdVisita(idVisita) { return getFatoVisitaByIdVisita(idVisita); }
function VS_listFatoVisitas() { return listFatoVisitas(); }
function VS_upsertAvaliacaoByVisitaCliente(obj) { return upsertAvaliacaoByVisitaCliente(obj); }
function VS_listAvaliacoesByVisita(idVisita) { return listAvaliacoesByVisita(idVisita); }
function VS_deleteAvaliacaoById(idAvaliacao) { return deleteAvaliacaoById(idAvaliacao); }

/* ===============================
   Util: distinct values (datalists)
================================ */

/**
 * Retorna valores distintos por coluna (para alimentar datalists)
 */
function getDistinctValues(sheetName, columns) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Aba "' + sheetName + '" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) {
    const out = {};
    (columns || []).forEach(c => out[c] = []);
    return out;
  }

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[h] = i; });

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const out = {};

  (columns || []).forEach(col => {
    const i = idx[col];
    const set = new Set();
    if (i === undefined) {
      out[col] = [];
      return;
    }
    data.forEach(r => {
      const v = String(r[i] ?? "").trim();
      if (v) set.add(v);
    });
    out[col] = Array.from(set).sort((a, b) => a.localeCompare(b));
  });

  return out;
}

/* ===============================
   Base_Clientes - abordagem robusta v2
   (não depende do DataService.listRecords)
================================ */

function _normHeader_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

/**
 * ✅ Lista mínima da Base_Clientes para o painel "Registros"
 * Retorna [{id, label}] com: ID, Nome Completo, Telefone, Email
 * - tolera cabeçalhos: "ID", "Id", "id", espaços, acentos etc.
 */
function BC_listForPanel_v2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Base_Clientes");
  if (!sh) throw new Error('Aba "Base_Clientes" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);
    if (key) normToIdx[key] = i;
  });

  const idxID = normToIdx["id"];
  const idxNome = normToIdx[_normHeader_("Nome Completo")];
  const idxTel = normToIdx[_normHeader_("Telefone")];
  const idxEmail = normToIdx[_normHeader_("Email")];

  if (idxID === undefined) throw new Error('Base_Clientes: coluna "ID" não encontrada (nem variação).');

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const out = [];

  data.forEach(r => {
    const id = r[idxID];
    if (id === "" || id === null || id === undefined) return;

    const nome = idxNome !== undefined ? String(r[idxNome] ?? "").trim() : "";
    let tel = idxTel !== undefined ? r[idxTel] : "";
    const email = idxEmail !== undefined ? String(r[idxEmail] ?? "").trim() : "";

    if (typeof tel === "number") tel = String(Math.trunc(tel));
    tel = String(tel ?? "").trim();

    out.push({
      id: String(id).trim(),
      label: [nome || "(sem nome)", tel || "(sem tel)", email || "(sem email)"].join(" • ")
    });
  });

  out.sort((a, b) => {
    const na = Number(a.id), nb = Number(b.id);
    if (!isNaN(na) && !isNaN(nb)) return nb - na;
    return String(b.id).localeCompare(String(a.id));
  });

  return out;
}

/**
 * Base_Clientes - lista para painel com filtros
 * filtros suportados: "Prazo de Compra" e "Status Atual"
 */
function BC_listForPanelFiltered_v1(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Base_Clientes");
  if (!sh) throw new Error('Aba "Base_Clientes" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);
    if (key) normToIdx[key] = i;
  });

  const idxID = normToIdx["id"];
  const idxNome = normToIdx[_normHeader_("Nome Completo")];
  const idxTel = normToIdx[_normHeader_("Telefone")];
  const idxEmail = normToIdx[_normHeader_("Email")];
  const idxPrazo = normToIdx[_normHeader_("Prazo de Compra")];
  const idxStatus = normToIdx[_normHeader_("Status Atual")];

  if (idxID === undefined) throw new Error('Base_Clientes: coluna "ID" não encontrada (nem variação).');

  const f = filters || {};
  const fPrazo = String(f.prazoCompra || "").trim().toLowerCase();
  const fStatus = String(f.statusAtual || "").trim().toLowerCase();

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const out = [];

  data.forEach(r => {
    const id = r[idxID];
    if (id === "" || id === null || id === undefined) return;

    const prazo = idxPrazo !== undefined ? String(r[idxPrazo] ?? "").trim() : "";
    const status = idxStatus !== undefined ? String(r[idxStatus] ?? "").trim() : "";

    if (fPrazo && prazo.toLowerCase() !== fPrazo) return;
    if (fStatus && status.toLowerCase() !== fStatus) return;

    const nome = idxNome !== undefined ? String(r[idxNome] ?? "").trim() : "";
    let tel = idxTel !== undefined ? r[idxTel] : "";
    const email = idxEmail !== undefined ? String(r[idxEmail] ?? "").trim() : "";

    if (typeof tel === "number") tel = String(Math.trunc(tel));
    tel = String(tel ?? "").trim();

    out.push({
      id: String(id).trim(),
      label: [nome || "(sem nome)", tel || "(sem tel)", email || "(sem email)"].join(" • ")
    });
  });

  out.sort((a, b) => {
    const na = Number(a.id), nb = Number(b.id);
    if (!isNaN(na) && !isNaN(nb)) return nb - na;
    return String(b.id).localeCompare(String(a.id));
  });

  return out;
}

/**
 * ✅ Carrega Base_Clientes por ID (direto da planilha) e retorna objeto completo
 * - Formata datas (se forem Date) para DD/MM/AAAA
 * - Aniversário (se for Date) para DD/MM
 */
function BC_getById_v2(idVal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Base_Clientes");
  if (!sh) throw new Error('Aba "Base_Clientes" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return null;

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);
    if (key) normToIdx[key] = i;
  });

  const idxID = normToIdx["id"];
  if (idxID === undefined) throw new Error('Base_Clientes: coluna "ID" não encontrada (nem variação).');

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const target = String(idVal).trim();

  const isDateObj = (v) => Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime());
  const pad2 = (n) => String(n).padStart(2, "0");
  const fmtDDMMYYYY = (d) => `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
  const fmtDDMM = (d) => `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}`;

  for (const r of data) {
    const idHere = String(r[idxID] ?? "").trim();
    if (idHere !== target) continue;

    const obj = {};
    headers.forEach((h, i) => {
      const keyRaw = String(h || "").trim();
      if (!keyRaw) return;

      let v = r[i];

      if (_normHeader_(keyRaw) === "telefone" && typeof v === "number") {
        v = String(Math.trunc(v));
      }

      const keyNorm = _normHeader_(keyRaw);

      if (["data primeiro contato", "ultimo contato", "proximo follow-up"].includes(keyNorm) && isDateObj(v)) {
        v = fmtDDMMYYYY(v);
      }

      if (keyNorm === _normHeader_("Data de Aniversário") && isDateObj(v)) {
        v = fmtDDMM(v);
      }

      obj[keyRaw] = (v === null || v === undefined) ? "" : v;
    });

    return obj;
  }

  return null;
}

// ===============================
// Leads_Compradores - abordagem robusta v2
// ===============================

function LC_listForPanel_v2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Leads_Compradores");
  if (!sh) throw new Error('Aba "Leads_Compradores" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);
    if (key) normToIdx[key] = i;
  });

  const idxTel = normToIdx[_normHeader_("Telefone")];
  if (idxTel === undefined) throw new Error('Leads_Compradores: coluna "Telefone" não encontrada.');

  const idxNome   = normToIdx[_normHeader_("Nome")];
  const idxOrc    = normToIdx[_normHeader_("Orçamento")];
  const idxTipo   = normToIdx[_normHeader_("Tipo Imóvel")];
  const idxBairro = normToIdx[_normHeader_("Bairro")];

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const out = [];

  data.forEach(r => {
    let tel = r[idxTel];
    if (tel === "" || tel === null || tel === undefined) return;

    if (typeof tel === "number") tel = String(Math.trunc(tel));
    tel = String(tel).trim();

    const nome   = idxNome   !== undefined ? String(r[idxNome]   ?? "").trim() : "";
    const orc    = idxOrc    !== undefined ? String(r[idxOrc]    ?? "").trim() : "";
    const tipo   = idxTipo   !== undefined ? String(r[idxTipo]   ?? "").trim() : "";
    const bairro = idxBairro !== undefined ? String(r[idxBairro] ?? "").trim() : "";

    // ✅ label exatamente como você pediu
    const label = [
      nome   || "(sem nome)",
      orc    || "(sem orçamento)",
      tipo   || "(sem tipo)",
      bairro || "(sem bairro)"
    ].join(" • ");

    out.push({ id: tel, label });
  });

  // ordena por nome (melhor UX)
  out.sort((a, b) => a.label.localeCompare(b.label, "pt-BR"));

  return out;
}

function LC_getById_v2(telefoneVal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Leads_Compradores");
  if (!sh) throw new Error('Aba "Leads_Compradores" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return null;

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);
    if (key) normToIdx[key] = i;
  });

  const idxTel = normToIdx[_normHeader_("Telefone")];
  if (idxTel === undefined) throw new Error('Leads_Compradores: coluna "Telefone" não encontrada.');

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();

  const normalizePhoneKey_ = (v) => {
    let raw = v;
    if (typeof raw === "number") raw = String(Math.trunc(raw));
    raw = String(raw ?? "").trim();
    const digits = raw.replace(/\D/g, "");
    return digits || raw;
  };

  const target = normalizePhoneKey_(telefoneVal);

  const isDateObj = (v) => Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime());
  const pad2 = (n) => String(n).padStart(2, "0");
  const fmtDDMMYYYY = (d) => `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;

  for (const r of data) {
    const tel = normalizePhoneKey_(r[idxTel]);
    if (tel !== target) continue;

    const obj = {};
    headers.forEach((h, i) => {
      const keyRaw = String(h || "").trim();
      if (!keyRaw) return;

      let v = r[i];

      // telefone number -> string
      if (_normHeader_(keyRaw) === "telefone") {
        v = normalizePhoneKey_(v);
      }

      // campos de data -> DD/MM/AAAA se vier Date
      const k = _normHeader_(keyRaw);
      if (["data entrada", "ultimo contato", "proximo follow-up"].includes(k) && isDateObj(v)) {
        v = fmtDDMMYYYY(v);
      }

      obj[keyRaw] = (v === null || v === undefined) ? "" : v;
    });

    return obj;
  }

  return null;
}

// ===============================
// Form_LeadsVendedores - funções exclusivas (LVD_*)
// ===============================

function LVD_listForPanel_v2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Leads_Vendedores");
  if (!sh) throw new Error('Aba "Leads_Vendedores" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);       // usa seu normalizador já existente no Code.gs
    if (key) normToIdx[key] = i;
  });

  const idxTel = normToIdx[_normHeader_("Telefone")];
  if (idxTel === undefined) throw new Error('Leads_Vendedores: coluna "Telefone" não encontrada.');

  const idxNome = normToIdx[_normHeader_("Nome Proprietário")];
  const idxQuadra = normToIdx[_normHeader_("Quadra/Endereço")];
  const idxStatus = normToIdx[_normHeader_("Status")];

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const out = [];

  data.forEach(r => {
    let tel = r[idxTel];
    if (tel === "" || tel === null || tel === undefined) return;

    if (typeof tel === "number") tel = String(Math.trunc(tel));
    tel = String(tel).trim();

    const nome = idxNome !== undefined ? String(r[idxNome] ?? "").trim() : "";
    const quadra = idxQuadra !== undefined ? String(r[idxQuadra] ?? "").trim() : "";
    const status = idxStatus !== undefined ? String(r[idxStatus] ?? "").trim() : "";

    const label = [
      nome || "(sem nome)",
      quadra || "(sem endereço)",
      status || "(sem status)"
    ].join(" • ");

    out.push({ id: tel, label });
  });

  out.sort((a, b) => a.label.localeCompare(b.label, "pt-BR"));
  return out;
}

/**
 * Form_LeadsVendedores - lista com filtro por Status
 */
function LVD_listForPanelFiltered_v1(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Leads_Vendedores");
  if (!sh) throw new Error('Aba "Leads_Vendedores" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);
    if (key) normToIdx[key] = i;
  });

  const idxTel = normToIdx[_normHeader_("Telefone")];
  const idxNome = normToIdx[_normHeader_("Nome Proprietário")];
  const idxQuadra = normToIdx[_normHeader_("Quadra/Endereço")];
  const idxStatus = normToIdx[_normHeader_("Status")];
  if (idxTel === undefined) throw new Error('Leads_Vendedores: coluna "Telefone" não encontrada.');

  const f = filters || {};
  const statusFilter = String(f.status || "").trim().toLowerCase();

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const out = [];

  data.forEach(r => {
    let tel = r[idxTel];
    if (tel === "" || tel === null || tel === undefined) return;

    if (typeof tel === "number") tel = String(Math.trunc(tel));
    tel = String(tel).trim();

    const nome = idxNome !== undefined ? String(r[idxNome] ?? "").trim() : "";
    const quadra = idxQuadra !== undefined ? String(r[idxQuadra] ?? "").trim() : "";
    const status = idxStatus !== undefined ? String(r[idxStatus] ?? "").trim() : "";

    if (statusFilter && status.toLowerCase() !== statusFilter) return;

    const label = [
      nome || "(sem nome)",
      quadra || "(sem endereço)",
      status || "(sem status)"
    ].join(" • ");

    out.push({ id: tel, label });
  });

  out.sort((a, b) => a.label.localeCompare(b.label, "pt-BR"));
  return out;
}

function LVD_getById_v2(telefoneVal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Leads_Vendedores");
  if (!sh) throw new Error('Aba "Leads_Vendedores" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return null;

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);
    if (key) normToIdx[key] = i;
  });

  const idxTel = normToIdx[_normHeader_("Telefone")];
  if (idxTel === undefined) throw new Error('Leads_Vendedores: coluna "Telefone" não encontrada.');

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const target = String(telefoneVal ?? "").trim();

  const isDateObj = (v) => Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime());
  const pad2 = (n) => String(n).padStart(2, "0");
  const fmtDDMMYYYY = (d) => `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;

  for (const r of data) {
    let tel = r[idxTel];
    if (typeof tel === "number") tel = String(Math.trunc(tel));
    tel = String(tel ?? "").trim();
    if (tel !== target) continue;

    const obj = {};
    headers.forEach((h, i) => {
      const keyRaw = String(h || "").trim();
      if (!keyRaw) return;

      let v = r[i];

      if (_normHeader_(keyRaw) === "telefone" && typeof v === "number") {
        v = String(Math.trunc(v));
      }

      const k = _normHeader_(keyRaw);
      if (["data entrada", "ultimo contato", "proximo follow-up"].includes(k) && isDateObj(v)) {
        v = fmtDDMMYYYY(v);
      }

      obj[keyRaw] = (v === null || v === undefined) ? "" : v;
    });

    return obj;
  }

  return null;
}

// ===============================
// Form_AgendarVisita - funções exclusivas (AG_*)
// (corrige hora "fixa" usando getDisplayValues)
// ===============================

function AG_listForPanel_v2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Agenda_Visitas");
  if (!sh) throw new Error('Aba "Agenda_Visitas" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);
    if (key) normToIdx[key] = i;
  });

  // Aceita "ID" ou "id"
  let idxId = normToIdx["id"];
  if (idxId === undefined) idxId = headers.findIndex(h => _normHeader_(h) === "id");
  if (idxId === -1 || idxId === undefined) throw new Error('Agenda_Visitas: coluna "ID/id" não encontrada.');

  const idxData = normToIdx[_normHeader_("Data")];
  const idxHora = normToIdx[_normHeader_("Hora")];
  const idxCli  = normToIdx[_normHeader_("Cliente")];
  const idxImv  = normToIdx[_normHeader_("Imóvel (Código)")];

  // ✅ pega o que o usuário vê na planilha (evita bug de timezone/serial em Hora)
  const range = sh.getRange(2, 1, lr - 1, lc);
  const values = range.getValues();           // para ID e eventuais checks
  const disp   = range.getDisplayValues();    // para Data/Hora/Cliente/Imóvel

  const out = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowD = disp[i];

    const id = row[idxId];
    if (id === "" || id === null || id === undefined) continue;

    const vData = (idxData !== undefined) ? String(rowD[idxData] ?? "").trim() : "";
    const vHora = (idxHora !== undefined) ? String(rowD[idxHora] ?? "").trim() : "";
    const vCli  = (idxCli  !== undefined) ? String(rowD[idxCli]  ?? "").trim() : "";
    const vImv  = (idxImv  !== undefined) ? String(rowD[idxImv]  ?? "").trim() : "";

    const label = `${vData || "(sem data)"} ${vHora || ""} • Cliente ${vCli || "-"} • Imóvel ${vImv || "-"}`;
    out.push({ id: String(id).trim(), label });
  }

  // Ordena por ID desc se numérico
  out.sort((a, b) => {
    const na = Number(a.id), nb = Number(b.id);
    if (!isNaN(na) && !isNaN(nb)) return nb - na;
    return String(b.id).localeCompare(String(a.id));
  });

  return out;
}

function AG_getById_v2(idVal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Agenda_Visitas");
  if (!sh) throw new Error('Aba "Agenda_Visitas" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return null;

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const key = _normHeader_(h);
    if (key) normToIdx[key] = i;
  });

  let idxId = normToIdx["id"];
  if (idxId === undefined) idxId = headers.findIndex(h => _normHeader_(h) === "id");
  if (idxId === -1 || idxId === undefined) throw new Error('Agenda_Visitas: coluna "ID/id" não encontrada.');

  const target = String(idVal ?? "").trim();
  const range = sh.getRange(2, 1, lr - 1, lc);
  const values = range.getValues();
  const disp   = range.getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowD = disp[i];

    const idHere = String(row[idxId] ?? "").trim();
    if (idHere !== target) continue;

    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const keyRaw = String(headers[c] || "").trim();
      if (!keyRaw) continue;

      // ✅ para Data/Hora devolve o DISPLAY (igual aparece na planilha)
      const k = _normHeader_(keyRaw);
      if (k === _normHeader_("data") || k === _normHeader_("hora")) {
        obj[keyRaw] = String(rowD[c] ?? "").trim();
      } else {
        const v = row[c];
        obj[keyRaw] = (v === null || v === undefined) ? "" : v;
      }
    }

    return obj;
  }

  return null;
}

// ======================================================
// Form_RegistrarVisita - funções exclusivas (RV_*)
// - Id_Visita auto incremental + DRAFT
// - Vínculo com Agenda salva em Id_Agendamento (Fato_Visitas)
// - Painéis/Selects usando getDisplayValues (hora/data confiáveis)
// ======================================================

function RV_ensureSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Fato_Visitas");
  if (!sh) throw new Error('Aba "Fato_Visitas" não existe.');

  const lc = sh.getLastColumn();
  if (lc < 1) throw new Error('Aba "Fato_Visitas" sem cabeçalho.');

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
  const norm = headers.map(h => _normHeader_(h));
  const needCols = ["id_visita", "id_agendamento"];

  // garante Id_Visita
  if (!norm.includes("id_visita")) {
    sh.insertColumnAfter(lc);
    sh.getRange(1, lc + 1).setValue("Id_Visita");
  }

  // recalcula após possível insert
  const lc2 = sh.getLastColumn();
  const headers2 = sh.getRange(1, 1, 1, lc2).getValues()[0].map(h => String(h || "").trim());
  const norm2 = headers2.map(h => _normHeader_(h));

  // garante Id_Agendamento
  if (!norm2.includes("id_agendamento")) {
    sh.insertColumnAfter(lc2);
    sh.getRange(1, lc2 + 1).setValue("Id_Agendamento");
  }

  // bold header
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight("bold");
}

/**
 * Reserva Id_Visita auto-incremental e cria um rascunho (DRAFT) em Fato_Visitas.
 * Isso permite que Fato_Avaliacao salve antes do "Salvar Visita" final.
 */
function RV_reserveIdVisita_v2() {
  RV_ensureSchema_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Fato_Visitas");
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[_normHeader_(h)] = i; });

  const idxIdVis = idx["id_visita"];
  if (idxIdVis === undefined) throw new Error('Fato_Visitas: coluna "Id_Visita" não encontrada.');

  // calcula próximo id
  let maxId = 0;
  if (lr >= 2) {
    const vals = sh.getRange(2, idxIdVis + 1, lr - 1, 1).getValues();
    vals.forEach(r => {
      const n = Number(String(r[0] ?? "").trim());
      if (!isNaN(n) && isFinite(n)) maxId = Math.max(maxId, n);
    });
  }
  const nextId = String(maxId + 1);

  // cria linha DRAFT (apenas Id_Visita preenchido; resto vazio)
  const row = new Array(lc).fill("");
  row[idxIdVis] = nextId;

  // se existir alguma coluna "Status" ou "status" colocamos DRAFT (não obrigatório)
  const idxStatus = idx["status"];
  if (idxStatus !== undefined) row[idxStatus] = "DRAFT";

  sh.appendRow(row);

  return { id_visita: nextId };
}

/**
 * Lista Agenda_Visitas para dropdown (id + label com Data/Hora/Cliente/Imóvel)
 * Usa getDisplayValues para Hora/Data não “virarem 19:23”.
 */
function RV_listAgendaForDropdown_v2() {
  // garante schema da agenda (seu ensureSchema_ já faz id)
  try { ensureSchema_(); } catch (e) {}

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Agenda_Visitas");
  if (!sh) return [];

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => { const k = _normHeader_(h); if (k) normToIdx[k] = i; });

  // id pode ser "id" ou "ID"
  let idxId = normToIdx["id"];
  if (idxId === undefined) idxId = headers.findIndex(h => _normHeader_(h) === "id");
  if (idxId === -1 || idxId === undefined) throw new Error('Agenda_Visitas: coluna "id/ID" não encontrada.');

  const idxData = normToIdx[_normHeader_("Data")];
  const idxHora = normToIdx[_normHeader_("Hora")];
  const idxCli  = normToIdx[_normHeader_("Cliente")];
  const idxImv  = normToIdx[_normHeader_("Imóvel (Código)")];

  const range = sh.getRange(2, 1, lr - 1, lc);
  const values = range.getValues();
  const disp   = range.getDisplayValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const id = values[i][idxId];
    if (id === "" || id === null || id === undefined) continue;

    const d = idxData !== undefined ? String(disp[i][idxData] ?? "").trim() : "";
    const h = idxHora !== undefined ? String(disp[i][idxHora] ?? "").trim() : "";
    const c = idxCli  !== undefined ? String(disp[i][idxCli]  ?? "").trim() : "";
    const m = idxImv  !== undefined ? String(disp[i][idxImv]  ?? "").trim() : "";

    const label = `${d || "(sem data)"} ${h || ""} • ${c ? ("Cliente " + c) : "Cliente -"} • Imóvel ${m || "-"}`;
    out.push({ id: String(id).trim(), label });
  }

  // id desc (se numérico)
  out.sort((a, b) => {
    const na = Number(a.id), nb = Number(b.id);
    if (!isNaN(na) && !isNaN(nb)) return nb - na;
    return String(b.id).localeCompare(String(a.id));
  });

  return out;
}

/**
 * Lê uma Agenda por id (retorna display de Data/Hora).
 */
function RV_getAgendaById_v2(idVal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Agenda_Visitas");
  if (!sh) return null;

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return null;

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => { const k = _normHeader_(h); if (k) normToIdx[k] = i; });

  let idxId = normToIdx["id"];
  if (idxId === undefined) idxId = headers.findIndex(h => _normHeader_(h) === "id");
  if (idxId === -1 || idxId === undefined) throw new Error('Agenda_Visitas: coluna "id/ID" não encontrada.');

  const target = String(idVal ?? "").trim();

  const range = sh.getRange(2, 1, lr - 1, lc);
  const values = range.getValues();
  const disp   = range.getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const idHere = String(values[i][idxId] ?? "").trim();
    if (idHere !== target) continue;

    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const keyRaw = String(headers[c] || "").trim();
      if (!keyRaw) continue;

      const k = _normHeader_(keyRaw);
      if (k === _normHeader_("data") || k === _normHeader_("hora")) obj[keyRaw] = String(disp[i][c] ?? "").trim();
      else obj[keyRaw] = (values[i][c] === null || values[i][c] === undefined) ? "" : values[i][c];
    }
    return obj;
  }

  return null;
}

/**
 * Cria agenda nova e retorna id.
 * (Usa a coluna id/ID já existente; se existir "id" usa ela; senão usa "ID")
 */
function RV_createAgenda_v2(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Agenda_Visitas");
  if (!sh) throw new Error('Aba "Agenda_Visitas" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => { const k = _normHeader_(h); if (k) normToIdx[k] = i; });

  let idxId = normToIdx["id"];
  if (idxId === undefined) idxId = headers.findIndex(h => _normHeader_(h) === "id");
  if (idxId === -1 || idxId === undefined) throw new Error('Agenda_Visitas: coluna "id/ID" não encontrada.');

  // próximo id (numérico)
  let maxId = 0;
  if (lr >= 2) {
    const vals = sh.getRange(2, idxId + 1, lr - 1, 1).getValues();
    vals.forEach(r => {
      const n = Number(String(r[0] ?? "").trim());
      if (!isNaN(n) && isFinite(n)) maxId = Math.max(maxId, n);
    });
  }
  const nextId = String(maxId + 1);

  const row = new Array(lc).fill("");
  row[idxId] = nextId;

  // preenche colunas conhecidas se existirem
  const map = {};
  headers.forEach((h, i) => { map[_normHeader_(h)] = i; });

  const setIf = (colName, value) => {
    const ix = map[_normHeader_(colName)];
    if (ix !== undefined) row[ix] = value;
  };

  setIf("Data", obj["Data"] || "");
  setIf("Hora", obj["Hora"] || "");
  setIf("Cliente", obj["Cliente"] || "");
  setIf("Imóvel (Código)", obj["Imóvel (Código)"] || "");
  setIf("Telefone", obj["Telefone"] || ""); // opcional

  sh.appendRow(row);

  return { id: nextId, action: "created" };
}


// ======================================================
// Form_RegistrarVisita - RV_* (ajustes finais)
// - Lista "Registros": só Id_Imovel + Data_Visita
// - Upload foto (Anexo_Ficha_Visita) para Drive com nome = Id_Visita
// ======================================================

function RV_getOrCreateFichaFolderId_() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty("FICHA_FOLDER_ID");
  if (folderId) {
    try { DriveApp.getFolderById(folderId); return folderId; } catch (e) {}
  }

  // cria pasta padrão (predeterminada pelo script property após criada)
  const root = DriveApp.getRootFolder();
  const name = "CRM_Corretagem_Fichas_Visita";
  const it = root.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : root.createFolder(name);

  folderId = folder.getId();
  props.setProperty("FICHA_FOLDER_ID", folderId);
  return folderId;
}

/**
 * Upload da foto da ficha de visita para Drive:
 * - salva em pasta fixa (Script Property FICHA_FOLDER_ID)
 * - nome do arquivo = "<Id_Visita>.<ext>"
 * - se já existir arquivo com esse nome, move p/ lixeira e substitui
 * Retorna {fileId, url, name}
 */
function RV_uploadFichaVisita_v2(idVisita, dataUrl) {
  const idv = String(idVisita || "").trim();
  if (!idv) throw new Error("Id_Visita inválido para upload.");

  const s = String(dataUrl || "");
  const m = s.match(/^data:(.+?);base64,(.+)$/);
  if (!m) throw new Error("Arquivo inválido (dataUrl).");

  const mime = m[1];
  const b64 = m[2];

  let ext = "jpg";
  if (mime === "image/png") ext = "png";
  else if (mime === "image/jpeg" || mime === "image/jpg") ext = "jpg";
  else if (mime === "image/webp") ext = "webp";
  else ext = "jpg";

  const bytes = Utilities.base64Decode(b64);
  const blob = Utilities.newBlob(bytes, mime, `${idv}.${ext}`);

  const folderId = RV_getOrCreateFichaFolderId_();
  const folder = DriveApp.getFolderById(folderId);

  // substitui arquivo anterior com mesmo nome
  const existing = folder.getFilesByName(blob.getName());
  while (existing.hasNext()) {
    const f = existing.next();
    f.setTrashed(true);
  }

  const file = folder.createFile(blob);

  // opcional: tentar deixar "qualquer pessoa com o link" (pode falhar se domínio restringe)
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  const url = file.getUrl();
  return { fileId: file.getId(), url, name: file.getName() };
}

/**
 * Lista Fato_Visitas para painel:
 * ✅ somente Id_Imovel e Data_Visita
 */
function RV_listFatosForPanel_v2() {
  RV_ensureSchema_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Fato_Visitas");
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => { const k = _normHeader_(h); if (k) normToIdx[k] = i; });

  const idxIdVis = normToIdx["id_visita"];
  const idxData  = normToIdx[_normHeader_("Data_Visita")];
  const idxImv   = normToIdx[_normHeader_("Id_Imovel")];
  const idxProp  = normToIdx[_normHeader_("Proposta")];
  if (idxIdVis === undefined) throw new Error('Fato_Visitas: coluna "Id_Visita" não encontrada.');

  const parseNum_ = (v) => {
    if (typeof v === "number") return isFinite(v) ? v : null;
    const s = String(v ?? "").trim();
    if (!s) return null;
    const raw = s.replace(/[R$\s]/g, "");
    let n = NaN;
    if (raw.includes(",")) {
      n = Number(raw.replace(/\./g, "").replace(",", "."));
    } else {
      n = Number(raw);
    }
    return isNaN(n) ? null : n;
  };

  const clientesMap = {};
  try {
    const shCli = ss.getSheetByName("Base_Clientes");
    if (shCli) {
      const lrCli = shCli.getLastRow();
      const lcCli = shCli.getLastColumn();
      if (lrCli >= 2 && lcCli >= 1) {
        const hCli = shCli.getRange(1, 1, 1, lcCli).getValues()[0];
        const mCli = {};
        hCli.forEach((h, i) => {
          const k = _normHeader_(h);
          if (k) mCli[k] = i;
        });
        const idxCliId = mCli["id"];
        const idxCliNome = mCli[_normHeader_("Nome Completo")];
        if (idxCliId !== undefined && idxCliNome !== undefined) {
          const dCli = shCli.getRange(2, 1, lrCli - 1, lcCli).getValues();
          dCli.forEach(r => {
            const id = String(r[idxCliId] ?? "").trim();
            const nome = String(r[idxCliNome] ?? "").trim();
            if (id) clientesMap[id] = nome || ("Cliente " + id);
          });
        }
      }
    }
  } catch (e) {}

  const aggByVisita = {};
  try {
    const shAv = ss.getSheetByName("Fato_Avaliacao");
    if (shAv) {
      const lrAv = shAv.getLastRow();
      const lcAv = shAv.getLastColumn();
      if (lrAv >= 2 && lcAv >= 1) {
        const hAv = shAv.getRange(1, 1, 1, lcAv).getValues()[0];
        const mAv = {};
        hAv.forEach((h, i) => {
          const k = _normHeader_(h);
          if (k) mAv[k] = i;
        });

        const idxAvVis = mAv[_normHeader_("Id_Visita")];
        const idxAvCli = mAv[_normHeader_("Id_Cliente")];
        const idxAvNota = mAv[_normHeader_("Nota_Geral")];
        const idxAvPreco = mAv[_normHeader_("Preco_N10")];

        if (idxAvVis !== undefined) {
          const dAv = shAv.getRange(2, 1, lrAv - 1, lcAv).getValues();
          dAv.forEach(r => {
            const idv = String(r[idxAvVis] ?? "").trim();
            if (!idv) return;

            if (!aggByVisita[idv]) {
              aggByVisita[idv] = {
                clientes: new Set(),
                sumNota: 0,
                cntNota: 0,
                sumPreco: 0,
                cntPreco: 0
              };
            }
            const a = aggByVisita[idv];

            if (idxAvCli !== undefined) {
              const idc = String(r[idxAvCli] ?? "").trim();
              if (idc) a.clientes.add(idc);
            }

            if (idxAvNota !== undefined) {
              const n = parseNum_(r[idxAvNota]);
              if (n !== null) { a.sumNota += n; a.cntNota++; }
            }

            if (idxAvPreco !== undefined) {
              const p = parseNum_(r[idxAvPreco]);
              if (p !== null) { a.sumPreco += p; a.cntPreco++; }
            }
          });
        }
      }
    }
  } catch (e) {}

  const range = sh.getRange(2, 1, lr - 1, lc);
  const values = range.getValues();
  const disp   = range.getDisplayValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const idv = values[i][idxIdVis];
    if (idv === "" || idv === null || idv === undefined) continue;

    const idVisita = String(idv).trim();
    const d  = idxData !== undefined ? String(disp[i][idxData] ?? "").trim() : "";
    const im = idxImv  !== undefined ? String(values[i][idxImv] ?? "").trim() : "";
    const proposta = idxProp !== undefined ? String(values[i][idxProp] ?? "").trim() : "";

    const agg = aggByVisita[idVisita] || null;
    const clientes = agg
      ? Array.from(agg.clientes).map(idc => clientesMap[idc] || ("Cliente " + idc))
      : [];

    const mediaNotaGeral = (agg && agg.cntNota > 0) ? (agg.sumNota / agg.cntNota) : null;
    const mediaPrecoN10 = (agg && agg.cntPreco > 0) ? (agg.sumPreco / agg.cntPreco) : null;

    const label = `Data ${d || "(sem data)"} • Imóvel ${im || "-"}`;
    out.push({
      id: idVisita,
      label,
      dataVisita: d,
      proposta,
      clientes,
      mediaNotaGeral,
      mediaPrecoN10
    });
  }

  out.sort((a, b) => Number(b.id) - Number(a.id));
  return out;
}

/**
 * Carrega Fato_Visitas por Id_Visita
 * (Data_Visita devolve display)
 */
function RV_getFatoByIdVisita_v2(idVisita) {
  RV_ensureSchema_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Fato_Visitas");
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return null;

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => { const k = _normHeader_(h); if (k) normToIdx[k] = i; });

  const idxIdVis = normToIdx["id_visita"];
  if (idxIdVis === undefined) throw new Error('Fato_Visitas: coluna "Id_Visita" não encontrada.');

  const target = String(idVisita ?? "").trim();

  const range = sh.getRange(2, 1, lr - 1, lc);
  const values = range.getValues();
  const disp   = range.getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const idHere = String(values[i][idxIdVis] ?? "").trim();
    if (idHere !== target) continue;

    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const keyRaw = String(headers[c] || "").trim();
      if (!keyRaw) continue;

      const k = _normHeader_(keyRaw);
      if (k === _normHeader_("data_visita")) obj[keyRaw] = String(disp[i][c] ?? "").trim();
      else obj[keyRaw] = (values[i][c] === null || values[i][c] === undefined) ? "" : values[i][c];
    }
    return obj;
  }

  return null;
}

/**
 * Salva (upsert) no rascunho DRAFT por Id_Visita.
 * - Se existir linha, atualiza
 * - Se não existir, cria
 */
function RV_upsertFatoVisita_v2(obj) {
  RV_ensureSchema_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Fato_Visitas");
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
  const normToIdx = {};
  headers.forEach((h, i) => { const k = _normHeader_(h); if (k) normToIdx[k] = i; });

  const idxIdVis = normToIdx["id_visita"];
  if (idxIdVis === undefined) throw new Error('Fato_Visitas: coluna "Id_Visita" não encontrada.');

  const idv = String(obj["Id_Visita"] ?? "").trim();
  if (!idv) throw new Error("Id_Visita é obrigatório.");

  // procura linha
  let rowIndex = -1;
  if (lr >= 2) {
    const vals = sh.getRange(2, idxIdVis + 1, lr - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0] ?? "").trim() === idv) { rowIndex = i + 2; break; }
    }
  }

  // monta linha alvo
  const writeRow = (rowIndex === -1) ? new Array(lc).fill("") : sh.getRange(rowIndex, 1, 1, lc).getValues()[0];

  // preenche colunas existentes pelo nome
  headers.forEach((h, i) => {
    if (!h) return;
    if (obj.hasOwnProperty(h)) writeRow[i] = obj[h];
  });

  // garante Id_Visita gravado
  writeRow[idxIdVis] = idv;

  if (rowIndex === -1) {
    sh.appendRow(writeRow);
    return { ok: true, action: "created", id: idv };
  } else {
    sh.getRange(rowIndex, 1, 1, lc).setValues([writeRow]);
    return { ok: true, action: "updated", id: idv };
  }
}

// ======================================================
// PDF Visitas - COMPLETO + compatibilidade com "Gerar_PDF_Visita"
// ======================================================

function PDF_getOrCreateFolderId_() {
  // Mantido por compatibilidade com versões antigas do projeto.
  // ⚠️ Não usar getRootFolder para evitar exigência de escopo adicional.
  return PDF_getFolderId_();
}

function PDF_getFolderId_() {
  const props = PropertiesService.getScriptProperties();
  const FALLBACK_FOLDER_ID = "1NfMPgTO6L_qSxFC3qn4CV9CIovNOJuls";
  const folderIdProp = String(props.getProperty("PDF_VISITAS_FOLDER_ID") || "").trim();
  const folderId = folderIdProp || FALLBACK_FOLDER_ID;

  try {
    DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(`Pasta inválida ou sem acesso. Verifique "PDF_VISITAS_FOLDER_ID" (atual: ${folderId}).`);
  }

  if (!folderIdProp && folderId) {
    try { props.setProperty("PDF_VISITAS_FOLDER_ID", folderId); } catch (e) {}
  }

  return folderId;
}

function _PDF_buildClientesPorVisita_() {
  const out = {};
  const cleanNome_ = (s) => String(s || "").trim().replace(/^cliente\s+/i, "").trim();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shA = ss.getSheetByName("Fato_Avaliacao");
    const shC = ss.getSheetByName("Base_Clientes");
    if (!shA || !shC) return out;

    const lrA = shA.getLastRow(), lcA = shA.getLastColumn();
    const lrC = shC.getLastRow(), lcC = shC.getLastColumn();
    if (lrA < 2 || lcA < 1 || lrC < 2 || lcC < 1) return out;

    const hA = shA.getRange(1, 1, 1, lcA).getValues()[0];
    const hC = shC.getRange(1, 1, 1, lcC).getValues()[0];
    const mA = {}, mC = {};
    hA.forEach((h, i) => { const k = _normHeader_(h); if (k) mA[k] = i; });
    hC.forEach((h, i) => { const k = _normHeader_(h); if (k) mC[k] = i; });

    const idxAVis = mA[_normHeader_("Id_Visita")];
    const idxACli = mA[_normHeader_("Id_Cliente")];
    const idxCID = mC["id"];
    const idxNome = mC[_normHeader_("Nome Completo")];
    if ([idxAVis, idxACli, idxCID, idxNome].some(x => x === undefined)) return out;

    const cliMap = {};
    shC.getRange(2, 1, lrC - 1, lcC).getValues().forEach(r => {
      const id = String(r[idxCID] ?? "").trim();
      if (!id) return;
      const nome = cleanNome_(r[idxNome]);
      cliMap[id] = nome || id;
    });

    shA.getRange(2, 1, lrA - 1, lcA).getValues().forEach(r => {
      const idv = String(r[idxAVis] ?? "").trim();
      const idc = String(r[idxACli] ?? "").trim();
      if (!idv || !idc) return;
      if (!out[idv]) out[idv] = new Set();
      out[idv].add(cliMap[idc] || idc);
    });
  } catch (e) {}
  return out;
}

function _PDF_getBackgroundDataUrl_() {
  const BG_FOLDER_ID = "1fAzFGRc4KCnY2ou-jPhQ9hoiauQj0-Ce";
  try {
    const folder = DriveApp.getFolderById(BG_FOLDER_ID);
    const it = folder.getFiles();
    let chosen = null;
    let chosenTime = 0;

    while (it.hasNext()) {
      const f = it.next();
      const name = String(f.getName() || "");
      if (!/\.(png|jpg|jpeg|webp)$/i.test(name)) continue;
      const t = (f.getLastUpdated() || f.getDateCreated()).getTime();
      if (t > chosenTime) {
        chosen = f;
        chosenTime = t;
      }
    }

    if (!chosen) return "";
    const blob = chosen.getBlob();
    const ct = blob.getContentType() || "image/png";
    const b64 = Utilities.base64Encode(blob.getBytes());
    return `data:${ct};base64,${b64}`;
  } catch (e) {
    return "";
  }
}

function PDF_listVisitasForSelect_v1() {
  RV_ensureSchema_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Fato_Visitas");
  if (!sh) throw new Error('Aba "Fato_Visitas" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => { const k = _normHeader_(h); if (k) normToIdx[k] = i; });

  const idxIdVis = normToIdx["id_visita"];
  const idxData  = normToIdx[_normHeader_("Data_Visita")];
  const idxImv   = normToIdx[_normHeader_("Id_Imovel")];
  const clientesPorVisita = _PDF_buildClientesPorVisita_();

  if (idxIdVis === undefined) throw new Error('Fato_Visitas: coluna "Id_Visita" não encontrada.');

  const range = sh.getRange(2, 1, lr - 1, lc);
  const values = range.getValues();
  const disp   = range.getDisplayValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const idv = values[i][idxIdVis];
    if (idv === "" || idv === null || idv === undefined) continue;

    const data = idxData !== undefined ? String(disp[i][idxData] ?? "").trim() : "";
    const imv  = idxImv  !== undefined ? String(values[i][idxImv] ?? "").trim() : "";

    const nomes = clientesPorVisita[String(idv).trim()]
      ? Array.from(clientesPorVisita[String(idv).trim()]).join(", ")
      : "(sem clientes)";
    const label = `${nomes} • Data: ${data || "(sem data)"}`;
    out.push({ id: String(idv).trim(), label });
  }

  out.sort((a, b) => Number(b.id) - Number(a.id));
  return out;
}

function PDF_getVisitaPayload_v1(idVisita) {
  const idv = String(idVisita || "").trim();
  if (!idv) return null; // ✅ não derruba

  const fato = RV_getFatoByIdVisita_v2(idv);
  if (!fato) throw new Error("Visita não encontrada em Fato_Visitas: " + idv);

  const dataVisita = String(fato["Data_Visita"] || "").trim();
  const idImovel = String(fato["Id_Imovel"] || "").trim();
  const idAgendamento = String(fato["Id_Agendamento"] || "").trim();
  const anexo = String(fato["Anexo_Ficha_Visita"] || "").trim();

  let imovel = null;
  try {
    if (idImovel) imovel = DataService.getById("Estoque_Imoveis", "Código", idImovel);
  } catch (e) { imovel = null; }

  let agenda = null;
  try {
    if (idAgendamento) agenda = RV_getAgendaById_v2(idAgendamento);
  } catch (e) { agenda = null; }

  let avaliacoes = [];
  try {
    avaliacoes = listAvaliacoesByVisita(idv) || [];
  } catch (e) { avaliacoes = []; }

  const clientesMap = {};
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shCli = ss.getSheetByName("Base_Clientes");
    if (shCli) {
      const lr = shCli.getLastRow();
      const lc = shCli.getLastColumn();
      if (lr >= 2 && lc >= 1) {
        const headers = shCli.getRange(1, 1, 1, lc).getValues()[0];
        const map = {};
        headers.forEach((h, i) => { const k = _normHeader_(h); if (k) map[k] = i; });
        const idxId = map["id"];
        const idxNome = map[_normHeader_("Nome Completo")];
        if (idxId !== undefined && idxNome !== undefined) {
          shCli.getRange(2, 1, lr - 1, lc).getValues().forEach(r => {
            const id = String(r[idxId] ?? "").trim();
            if (!id) return;
            const nome = String(r[idxNome] ?? "").trim().replace(/^cliente\s+/i, "").trim();
            clientesMap[id] = nome || id;
          });
        }
      }
    }
  } catch (e) {}

  avaliacoes = (avaliacoes || []).map(a => {
    const idc = String(a["Id_Cliente"] || "").trim();
    const nome = String(clientesMap[idc] || idc || "").replace(/^cliente\s+/i, "").trim();
    return { ...a, Cliente_Nome: nome };
  });

  const notasGerais = avaliacoes
    .map(a => Number(a["Nota_Geral"]))
    .filter(n => !isNaN(n));
  const notaMedia = notasGerais.length ? (notasGerais.reduce((s, n) => s + n, 0) / notasGerais.length) : null;

  const clientesNomes = Array.from(new Set(
    avaliacoes
      .map(a => String(a.Cliente_Nome || "").trim())
      .filter(Boolean)
  ));

  return {
    id_visita: idv,
    fato,
    data_visita: dataVisita,
    id_imovel: idImovel,
    id_agendamento: idAgendamento,
    anexo_ficha_visita: anexo,
    imovel,
    agenda,
    avaliacoes,
    nota_media: notaMedia,
    clientes_nomes: clientesNomes,
    bg_data_url: _PDF_getBackgroundDataUrl_()
  };
}

function PDF_generatePdfVisita_v1(idVisita) {
  const payload = PDF_getVisitaPayload_v1(idVisita);
  if (!payload) throw new Error("Parâmetro visitaId ausente ou visita não encontrada.");

  const tpl = HtmlService.createTemplateFromFile("Pdf_Visita_Template");
  tpl.data = payload;
  const html = tpl.evaluate().getContent();

  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF);
  const name = `Visita_${payload.id_visita}.pdf`;
  blob.setName(name);

  const folder = DriveApp.getFolderById(PDF_getFolderId_()); // ✅ sem root

  const existing = folder.getFilesByName(name);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const file = folder.createFile(blob);

  // se der erro de permissão aqui, não derruba o processo
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  return { ok: true, fileId: file.getId(), name: file.getName(), url: file.getUrl() };
}

/**
 * ✅ COMPATIBILIDADE TOTAL:
 * Aceita:
 * - Gerar_PDF_Visita("123")
 * - Gerar_PDF_Visita({visitaId:"123"})
 * - Gerar_PDF_Visita({idVisita:"123"})
 */
/**
 * ✅ SAFE: nunca derruba o WebApp durante template render
 * Aceita:
 * - Gerar_PDF_Visita("123")
 * - Gerar_PDF_Visita({visitaId:"123"})
 * - Gerar_PDF_Visita({idVisita:"123"})
 *
 * Retorna sempre {ok:boolean, ...}
 */
function Gerar_PDF_Visita(arg) {
  try {
    let visitaId = "";

    if (typeof arg === "string" || typeof arg === "number") {
      visitaId = String(arg).trim();
    } else if (arg && typeof arg === "object") {
      visitaId = String(arg.visitaId || arg.idVisita || "").trim();
    }

    if (!visitaId) {
      return { ok: false, message: "Parâmetro visitaId ausente." };
    }

    const res = PDF_generatePdfVisita_v1(visitaId); // esta pode lançar
    // força formato consistente
    return { ok: true, ...res };

  } catch (e) {
    return { ok: false, message: e && e.message ? e.message : String(e) };
  }
}

// ===============================
// Estoque_Imoveis - robusto v2 (novo schema)
// ===============================

function _EST_getHeaders_(sh) {
  const lc = sh.getLastColumn();
  if (lc < 1) return [];
  return sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
}

function _EST_norm_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "");
}

function _EST_findCol_(headers, wanted) {
  // wanted: array de alternativas
  const map = {};
  headers.forEach((h, i) => { map[_EST_norm_(h)] = i; });
  for (const w of wanted) {
    const ix = map[_EST_norm_(w)];
    if (ix !== undefined) return ix;
  }
  return undefined;
}

/**
 * ✅ Dropdown de imóveis para selects do sistema
 * Retorna [{id,label}] onde:
 * - id = Codigo (novo) (ou Código antigo, se existir)
 * - label = "Codigo • Tipo • Valor • Bairro"
 */
function listImoveisForSelect() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Estoque_Imoveis");
  if (!sh) return [];

  const lr = sh.getLastRow();
  const headers = _EST_getHeaders_(sh);
  const lc = headers.length;
  if (lr < 2 || lc < 1) return [];

  const idxCodigo  = _EST_findCol_(headers, ["Codigo", "Código"]);
  const idxTipo    = _EST_findCol_(headers, ["Tipo"]);
  const idxValor   = _EST_findCol_(headers, ["Valor", "Preço", "Preco"]);
  const idxBairro  = _EST_findCol_(headers, ["Bairro"]);

  if (idxCodigo === undefined) throw new Error('Estoque_Imoveis: coluna "Codigo" não encontrada.');

  const range = sh.getRange(2, 1, lr - 1, lc);
  const values = range.getValues();
  const disp   = range.getDisplayValues(); // mostra Valor formatado, se houver

  const out = [];

  for (let i = 0; i < values.length; i++) {
    const codigo = String(values[i][idxCodigo] ?? "").trim();
    if (!codigo) continue;

    const tipo   = (idxTipo   !== undefined) ? String(values[i][idxTipo] ?? "").trim() : "";
    const valor  = (idxValor  !== undefined) ? String(disp[i][idxValor]  ?? "").trim() : "";
    const bairro = (idxBairro !== undefined) ? String(values[i][idxBairro] ?? "").trim() : "";

    const label = `${codigo} • ${tipo || "-"} • ${valor || "-"} • ${bairro || "-"}`;
    out.push({ id: codigo, label });
  }

  out.sort((a, b) => a.label.localeCompare(b.label, "pt-BR"));
  return out;
}

// ======================================================
// ESTOQUE - Import XLSX + Explorador com filtros (v1)
// Requer: Advanced Drive Service habilitado (Drive API)
// ======================================================

const EST_CFG = {
  SHEET: "Estoque_Imoveis",
  HEADERS: ["Codigo","Captadores","Tipo","Quartos","Valor","Endereco","Bairro","PublicacaoNaInternet","Exclusivo"]
};

function EST_getDistinctFilters_v1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(EST_CFG.SHEET);
  if (!sh) throw new Error(`Aba "${EST_CFG.SHEET}" não existe.`);

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return { Bairro: [], Quartos: [] };

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
  const idxB = headers.indexOf("Bairro");
  const idxQ = headers.indexOf("Quartos");

  const out = { Bairro: [], Quartos: [] };
  if (idxB === -1 && idxQ === -1) return out;

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const setB = new Set();
  const setQ = new Set();

  data.forEach(r => {
    if (idxB !== -1) {
      const v = String(r[idxB] ?? "").trim();
      if (v) setB.add(v);
    }
    if (idxQ !== -1) {
      const v = String(r[idxQ] ?? "").trim();
      if (v) setQ.add(v);
    }
  });

  out.Bairro = Array.from(setB).sort((a,b)=>a.localeCompare(b,"pt-BR"));
  out.Quartos = Array.from(setQ).sort((a,b)=>a.localeCompare(b,"pt-BR"));
  return out;
}

/**
 * Lista do painel Registros com filtros combináveis:
 * filters = { bairro?:string, quartos?:string, vmin?:number|null, vmax?:number|null }
 * Retorna [{id,label,raw}]
 */
function EST_listForPanelFiltered_v1(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(EST_CFG.SHEET);
  if (!sh) return [];

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
  const idx = {};
  headers.forEach((h,i)=>{ if(h) idx[h]=i; });

  // ✅ inclui Captadores
  const need = ["Codigo","Bairro","Quartos","Valor","Tipo","Captadores"];
  need.forEach(k=>{
    if (idx[k] === undefined) throw new Error(`Estoque_Imoveis: coluna "${k}" não encontrada.`);
  });

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  const disp = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();

  const fbairro = String(filters?.bairro || "").trim();
  const fquartos = String(filters?.quartos || "").trim();
  const vmin = (filters && typeof filters.vmin === "number") ? filters.vmin : null;
  const vmax = (filters && typeof filters.vmax === "number") ? filters.vmax : null;

  function toNumBRL_(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return v;
    const s = String(v).trim();
    if (!s) return null;
    const n = Number(s.replace(/[^\d,.-]/g,"").replace(/\./g,"").replace(",","."));
    return isNaN(n) ? null : n;
  }

  const out = [];

  for (let i=0;i<data.length;i++){
    const codigo = String(data[i][idx["Codigo"]] ?? "").trim();
    if (!codigo) continue;

    const captadores = String(data[i][idx["Captadores"]] ?? "").trim();
    const bairro = String(data[i][idx["Bairro"]] ?? "").trim();
    const quartos = String(data[i][idx["Quartos"]] ?? "").trim();
    const tipo = String(data[i][idx["Tipo"]] ?? "").trim();

    const valorNum = toNumBRL_(data[i][idx["Valor"]]);
    const valorDisp = String(disp[i][idx["Valor"]] ?? "").trim();

    if (fbairro && bairro !== fbairro) continue;
    if (fquartos && quartos !== fquartos) continue;
    if (vmin !== null && (valorNum === null || valorNum < vmin)) continue;
    if (vmax !== null && (valorNum === null || valorNum > vmax)) continue;

    out.push({
      id: codigo, // ✅ continua sendo o ID real
      // ✅ tira o codigo daqui para não duplicar
      label: `${captadores || "-"} • ${tipo || "-"} • ${quartos || "-"}Q • ${bairro || "-"} • ${valorDisp || "-"}`,
      raw: {
        Codigo: codigo,
        Captadores: captadores,
        Tipo: tipo,
        Quartos: quartos,
        Bairro: bairro,
        Valor: valorDisp
      }
    });
  }

  out.sort((a,b)=>a.label.localeCompare(b.label,"pt-BR"));
  return out;
}

/**
 * Importa XLSX (base64) e substitui completamente a aba Estoque_Imoveis.
 * - Exige que o XLSX tenha EXATAMENTE as colunas EST_CFG.HEADERS (na primeira linha).
 * - Apaga registros antigos (mantém header).
 *
 * Requer Advanced Drive Service (Drive) ON.
 */
function EST_importXlsxReplace_v1(fileBase64, fileName) {
  if (!fileBase64) throw new Error("Arquivo vazio.");
  fileName = String(fileName || "estoque.xlsx");

  // Blob do upload
  const bytes = Utilities.base64Decode(fileBase64);
  const blob = Utilities.newBlob(bytes, MimeType.MICROSOFT_EXCEL, fileName);

  // 1) cria arquivo temporário no Drive
  const tmp = DriveApp.createFile(blob);
  const tmpId = tmp.getId();

  try {
    // 2) converte para Google Sheets (Drive Advanced Service)
    // ATENÇÃO: precisa habilitar Drive API no projeto
    const resource = {
      title: "TMP_IMPORT_ESTOQUE_" + new Date().toISOString(),
      mimeType: MimeType.GOOGLE_SHEETS
    };

    const converted = Drive.Files.copy(resource, tmpId, { convert: true });
    const gsId = converted.id;

    // 3) lê a planilha convertida
    const imp = SpreadsheetApp.openById(gsId);
    const impSh = imp.getSheets()[0];

    const lr = impSh.getLastRow();
    const lc = impSh.getLastColumn();
    if (lr < 2 || lc < 1) throw new Error("XLSX sem dados (mínimo: header + 1 linha).");

    const headers = impSh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());

    // valida header exato (mesma ordem e nomes)
    const expected = EST_CFG.HEADERS;
    const sameLen = headers.length >= expected.length;
    if (!sameLen) throw new Error("Header do XLSX inválido: colunas insuficientes.");

    for (let i=0; i<expected.length; i++){
      if (headers[i] !== expected[i]) {
        throw new Error(`Header do XLSX inválido na coluna ${i+1}. Esperado "${expected[i]}", veio "${headers[i]}".`);
      }
    }

    // pega somente as colunas esperadas
    const values = impSh.getRange(2, 1, lr - 1, expected.length).getValues();

    // 4) escreve na aba destino (substitui tudo)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dest = ss.getSheetByName(EST_CFG.SHEET) || ss.insertSheet(EST_CFG.SHEET);

    // garante header correto
    dest.clearContents();
    dest.getRange(1, 1, 1, expected.length).setValues([expected]).setFontWeight("bold");

    if (values.length) {
      dest.getRange(2, 1, values.length, expected.length).setValues(values);
    }

    // formata Valor como moeda (opcional)
    const idxValor = expected.indexOf("Valor");
    if (idxValor !== -1 && values.length) {
      dest.getRange(2, idxValor + 1, values.length, 1).setNumberFormat('"R$" #,##0.00');
    }

    return { ok: true, rows: values.length };

  } catch (e) {
    // dica específica se Drive API não estiver habilitada
    const msg = String(e && e.message ? e.message : e);
    if (msg.includes("Drive") || msg.includes("Files")) {
      throw new Error(
        msg +
        "\n\n⚠️ Para importar XLSX, habilite: Apps Script → Services → Advanced Google services → Drive API (ON) e no GCP Console habilite Drive API."
      );
    }
    throw e;

  } finally {
    // limpa temporários (tenta)
    try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {}
  }
}


/***********************
 * Leads_Compradores - filtros + panel (v4)
 * - filtros: Bairro, Tipo Imóvel, Orçamento (dropdown), Status (dropdown), Data Entrada (de/até)
 * - panel label: Nome • Telefone • Tipo Imóvel
 ***********************/

function LC_getDistinctFilters_v4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Leads_Compradores");
  if (!sh) throw new Error('Aba "Leads_Compradores" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return { Bairro: [], "Tipo Imóvel": [], "Orçamento": [], "Status": [] };

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const k = _normHeader_(h);
    if (k) normToIdx[k] = i;
  });

  const idxBairro = normToIdx[_normHeader_("Bairro")];
  const idxTipo   = normToIdx[_normHeader_("Tipo Imóvel")];
  const idxOrc    = normToIdx[_normHeader_("Orçamento")];
  const idxStatus = normToIdx[_normHeader_("Status")];

  const setB = new Set(), setT = new Set(), setO = new Set(), setS = new Set();

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  data.forEach(r => {
    if (idxBairro !== undefined) {
      const v = String(r[idxBairro] ?? "").trim();
      if (v) setB.add(v);
    }
    if (idxTipo !== undefined) {
      const v = String(r[idxTipo] ?? "").trim();
      if (v) setT.add(v);
    }
    if (idxOrc !== undefined) {
      const v = String(r[idxOrc] ?? "").trim();
      if (v) setO.add(v);
    }
    if (idxStatus !== undefined) {
      const v = String(r[idxStatus] ?? "").trim();
      if (v) setS.add(v);
    }
  });

  return {
    Bairro: Array.from(setB).sort((a,b)=>a.localeCompare(b,"pt-BR")),
    "Tipo Imóvel": Array.from(setT).sort((a,b)=>a.localeCompare(b,"pt-BR")),
    "Orçamento": Array.from(setO).sort((a,b)=>a.localeCompare(b,"pt-BR")),
    "Status": Array.from(setS).sort((a,b)=>a.localeCompare(b,"pt-BR"))
  };
}

function LC_listForPanelFiltered_v4(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Leads_Compradores");
  if (!sh) throw new Error('Aba "Leads_Compradores" não existe.');

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0];
  const normToIdx = {};
  headers.forEach((h, i) => {
    const k = _normHeader_(h);
    if (k) normToIdx[k] = i;
  });

  const idxTel    = normToIdx[_normHeader_("Telefone")];
  const idxNome   = normToIdx[_normHeader_("Nome")];
  const idxTipo   = normToIdx[_normHeader_("Tipo Imóvel")];
  const idxBairro = normToIdx[_normHeader_("Bairro")];
  const idxOrc    = normToIdx[_normHeader_("Orçamento")];
  const idxStatus = normToIdx[_normHeader_("Status")];
  const idxDataE  = normToIdx[_normHeader_("Data Entrada")];

  if (idxTel === undefined) throw new Error('Leads_Compradores: coluna "Telefone" não encontrada.');

  const f = filters || {};
  const fBairro = String(f.bairro || "").trim();
  const fTipo   = String(f.tipoImovel || "").trim();
  const fOrc    = String(f.orcamento || "").trim();
  const fStatus = String(f.status || "").trim();

  const dtIni = (f.dataIni && String(f.dataIni).trim()) ? parseDDMMYYYY_(String(f.dataIni).trim()) : null;
  const dtFim = (f.dataFim && String(f.dataFim).trim()) ? parseDDMMYYYY_(String(f.dataFim).trim()) : null;

  // se usuário preencheu algo inválido, falha com mensagem boa
  if (f.dataIni && String(f.dataIni).trim() && !dtIni) throw new Error('Data inicial inválida. Use DD/MM/AAAA.');
  if (f.dataFim && String(f.dataFim).trim() && !dtFim) throw new Error('Data final inválida. Use DD/MM/AAAA.');

  const range = sh.getRange(2, 1, lr - 1, lc);
  const values = range.getValues();
  const disp   = range.getDisplayValues();

  const out = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowD= disp[i];

    // ID = telefone
    let tel = row[idxTel];
    if (tel === "" || tel === null || tel === undefined) continue;
    if (typeof tel === "number") tel = String(Math.trunc(tel));
    tel = String(tel).trim();

    const nome   = idxNome   !== undefined ? String(row[idxNome]   ?? "").trim() : "";
    const tipo   = idxTipo   !== undefined ? String(row[idxTipo]   ?? "").trim() : "";
    const bairro = idxBairro !== undefined ? String(row[idxBairro] ?? "").trim() : "";
    const orc    = idxOrc    !== undefined ? String(row[idxOrc]    ?? "").trim() : "";
    const status = idxStatus !== undefined ? String(row[idxStatus] ?? "").trim() : "";

    // filtros somáveis
    if (fBairro && bairro !== fBairro) continue;
    if (fTipo && tipo !== fTipo) continue;
    if (fOrc && orc !== fOrc) continue;
    if (fStatus && status !== fStatus) continue;

    // data entrada (de/até)
    if (dtIni || dtFim) {
      let rawDate = idxDataE !== undefined ? row[idxDataE] : null;

      // preferir Date real; se não, tentar display
      let d = parseAnyDate_(rawDate);

      if (!d && idxDataE !== undefined) {
        const dv = String(rowD[idxDataE] ?? "").trim();
        d = parseAnyDate_(dv);
      }

      if (!d) continue;
      if (dtIni && d < dtIni) continue;
      if (dtFim && d > dtFim) continue;
    }

    // label: Nome • Telefone • Tipo Imóvel
    const label = [nome || "(sem nome)", tel || "(sem tel)", tipo || "(sem tipo)"].join(" • ");
    out.push({ id: tel, label });
  }

  out.sort((a, b) => a.label.localeCompare(b.label, "pt-BR"));
  return out;
}
