/**
 * Gerar_PDF_Visita.js (ATUALIZADO PARA O CRM ATUAL)
 * - SEM doGet (evita conflito com o WebApp principal)
 * - Fonte de dados: Fato_Visitas, Fato_Avaliacao, Base_Clientes, Agenda_Visitas, Estoque_Imoveis
 * - Gera PDF via template HtmlService: Pdf_Visita_Template.html
 * - Pasta destino no Drive: ScriptProperties "PDF_VISITAS_FOLDER_ID" (auto-cria se não existir)
 *
 * Funções expostas para UI:
 *   - PDF_listVisitasForSelect_v1()
 *   - PDF_getVisitaPayload_v1(idVisita)
 *   - Gerar_PDF_Visita(arg)   // SAFE wrapper
 *   - PDF_generatePdfVisita_v1(idVisita)
 */

const PDF_CFG = {
  SHEETS: {
    FATO: "Fato_Visitas",
    AVAL: "Fato_Avaliacao",
    CLIENTES: "Base_Clientes",
    AGENDA: "Agenda_Visitas",
    ESTOQUE: "Estoque_Imoveis"
  },
  KEYS: {
    FATO_ID: "Id_Visita",
    FATO_IMOVEL: "Id_Imovel",
    FATO_DATA: "Data_Visita",
    FATO_ANEXO: "Anexo_Ficha_Visita",
    FATO_ID_AGENDAMENTO: "Id_Agendamento",

    AVAL_ID: "id_Avaliacao",
    AVAL_ID_VISITA: "Id_Visita",
    AVAL_ID_CLIENTE: "Id_Cliente",

    CLI_ID: "ID",
    CLI_NOME: "Nome Completo",
    CLI_TEL: "Telefone",
    CLI_EMAIL: "Email",

    AGENDA_ID: "ID",
    ESTOQUE_ID: "Código"
  }
};

/* =========================
   Pasta de destino do PDF
========================= */
function PDF_getFolderId_() {
  const props = PropertiesService.getScriptProperties();
  const FALLBACK_FOLDER_ID = "1NfMPgTO6L_qSxFC3qn4CV9CIovNOJuls";
  const folderIdProp = String(props.getProperty("PDF_VISITAS_FOLDER_ID") || "").trim();
  const folderId = folderIdProp || FALLBACK_FOLDER_ID;

  // valida se a pasta existe (isso já exige Drive scope — mas é a validação correta)
  try {
    DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(
      `Pasta inválida ou sem acesso. Verifique "PDF_VISITAS_FOLDER_ID" (atual: ${folderId}).`
    );
  }

  if (!folderIdProp && folderId) {
    try { props.setProperty("PDF_VISITAS_FOLDER_ID", folderId); } catch (e) {}
  }

  return folderId;
}
/* =========================
   Util: leitura robusta
   (sempre como texto / display)
========================= */
function _pdf_ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _pdf_sheet_(name) {
  const sh = _pdf_ss_().getSheetByName(name);
  if (!sh) throw new Error(`Aba "${name}" não existe.`);
  return sh;
}

function _pdf_headers_(sh) {
  const lc = sh.getLastColumn();
  if (lc < 1) throw new Error(`Aba "${sh.getName()}" sem cabeçalho.`);
  return sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(h => String(h || "").trim());
}

function _pdf_mapHeaders_(headers) {
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });
  return map;
}

function _pdf_getAllObjects_(sheetName) {
  const sh = _pdf_sheet_(sheetName);
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = _pdf_headers_(sh);
  const values = sh.getRange(2, 1, lr - 1, headers.length).getDisplayValues();

  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
    return obj;
  });
}

function _pdf_findByKey_(sheetName, keyCol, keyVal) {
  const key = String(keyVal || "").trim();
  if (!key) return null;

  const sh = _pdf_sheet_(sheetName);
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return null;

  const headers = _pdf_headers_(sh);
  const map = _pdf_mapHeaders_(headers);

  if (map[keyCol] === undefined) {
    throw new Error(`Aba "${sheetName}": coluna "${keyCol}" não encontrada.`);
  }

  const data = sh.getRange(2, 1, lr - 1, headers.length).getDisplayValues();
  const idx = map[keyCol];

  for (let i = 0; i < data.length; i++) {
    const v = String(data[i][idx] || "").trim();
    if (v === key) {
      const obj = {};
      headers.forEach((h, j) => { if (h) obj[h] = data[i][j]; });
      return obj;
    }
  }
  return null;
}

function _pdf_filterByKey_(sheetName, keyCol, keyVal) {
  const key = String(keyVal || "").trim();
  if (!key) return [];

  const sh = _pdf_sheet_(sheetName);
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = _pdf_headers_(sh);
  const map = _pdf_mapHeaders_(headers);

  if (map[keyCol] === undefined) {
    throw new Error(`Aba "${sheetName}": coluna "${keyCol}" não encontrada.`);
  }

  const data = sh.getRange(2, 1, lr - 1, headers.length).getDisplayValues();
  const idx = map[keyCol];

  const out = [];
  for (let i = 0; i < data.length; i++) {
    const v = String(data[i][idx] || "").trim();
    if (v === key) {
      const obj = {};
      headers.forEach((h, j) => { if (h) obj[h] = data[i][j]; });
      out.push(obj);
    }
  }
  return out;
}

/* =========================
   Util: nota média
========================= */
function _pdf_num_(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (!s) return null;
  // suporta "1.234,56" / "1234.56"
  const n = Number(s.replace(/[^\d,.-]/g, "").replace(/\./g, "").replace(",", "."));
  return isNaN(n) ? null : n;
}

function _pdf_calcNotaMedia_(avaliacoes) {
  if (!avaliacoes || !avaliacoes.length) return null;
  const vals = avaliacoes
    .map(a => _pdf_num_(a["Nota_Geral"]))
    .filter(n => n !== null);
  if (!vals.length) return null;
  const m = vals.reduce((acc, x) => acc + x, 0) / vals.length;
  return m;
}


function _pdf_getBackgroundDataUrl_() {
  const BG_FOLDER_ID = "1fAzFGRc4KCnY2ou-jPhQ9hoiauQj0-Ce";
  const SUPPORTED_MIME_TYPES = {
    "image/png": true,
    "image/jpeg": true,
    "image/jpg": true,
    "image/gif": true
  };
  try {
    const folder = DriveApp.getFolderById(BG_FOLDER_ID);
    const it = folder.getFiles();
    let chosen = null;
    let chosenTime = 0;

    while (it.hasNext()) {
      const f = it.next();
      let ct = "";
      try { ct = String(f.getMimeType() || "").toLowerCase(); } catch (e) { ct = ""; }
      if (!SUPPORTED_MIME_TYPES[ct]) continue;

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

function _pdf_buildClientesPorVisita_() {
  const S = PDF_CFG.SHEETS;
  const K = PDF_CFG.KEYS;
  const out = {};
  const cleanNome_ = (s) => String(s || "").trim().replace(/^cliente\s+/i, "").trim();

  const avs = _pdf_getAllObjects_(S.AVAL);
  const clientes = _pdf_getAllObjects_(S.CLIENTES);
  const cliMap = {};

  clientes.forEach(c => {
    const idc = String(c[K.CLI_ID] || "").trim();
    if (!idc) return;
    const nome = cleanNome_(c[K.CLI_NOME]);
    cliMap[idc] = nome || idc;
  });

  avs.forEach(a => {
    const idv = String(a[K.AVAL_ID_VISITA] || "").trim();
    const idc = String(a[K.AVAL_ID_CLIENTE] || "").trim();
    if (!idv || !idc) return;
    if (!out[idv]) out[idv] = new Set();
    out[idv].add(cliMap[idc] || idc);
  });

  return out;
}

/* =========================
   UI: lista p/ dropdown
   (Data_Visita + nomes dos clientes vinculados por Fato_Avaliacao)
========================= */
function PDF_listVisitasForSelect_v1() {
  const S = PDF_CFG.SHEETS;
  const K = PDF_CFG.KEYS;

  const fatos = _pdf_getAllObjects_(S.FATO);

  const clientesPorVisita = _pdf_buildClientesPorVisita_();

  const rows = fatos
    .map(f => {
      const idv = String(f[K.FATO_ID] || "").trim();
      if (!idv) return null;

      const data = String(f[K.FATO_DATA] || "").trim();
      const nomes = clientesPorVisita[idv] ? Array.from(clientesPorVisita[idv]).join(", ") : "(sem clientes)";

      return {
        id: idv,
        label: `${nomes} • Data: ${data || "-"}`
      };
    })
    .filter(Boolean);

  // ordena por Id_Visita desc (se for numérico)
  rows.sort((a, b) => {
    const na = Number(a.id), nb = Number(b.id);
    if (!isNaN(na) && !isNaN(nb)) return nb - na;
    return String(b.id).localeCompare(String(a.id));
  });

  return rows;
}

/* =========================
   Payload do PDF (modelo atual)
========================= */
function PDF_getVisitaPayload_v1(idVisita) {
  const S = PDF_CFG.SHEETS;
  const K = PDF_CFG.KEYS;

  const idv = String(idVisita || "").trim();
  if (!idv) return null; // ✅ não derruba web

  // Fato_Visitas (obrigatório)
  const fato = _pdf_findByKey_(S.FATO, K.FATO_ID, idv);
  if (!fato) return null;

  // Estoque_Imoveis (opcional)
  const idImovel = String(fato[K.FATO_IMOVEL] || "").trim();
  const imovel = idImovel ? _pdf_findByKey_(S.ESTOQUE, K.ESTOQUE_ID, idImovel) : null;

  // Agenda_Visitas (opcional)
  const idAg = String(fato[K.FATO_ID_AGENDAMENTO] || "").trim();
  const agenda = idAg ? _pdf_findByKey_(S.AGENDA, K.AGENDA_ID, idAg) : null;

  // Avaliações (0..n)
  const avs = _pdf_filterByKey_(S.AVAL, K.AVAL_ID_VISITA, idv);

  // Enriquecer com nome do cliente (Base_Clientes)
  // Construímos um map rápido ID->(nome,tel,email)
  const clientes = _pdf_getAllObjects_(S.CLIENTES);
  const cliMap = {};
  clientes.forEach(c => {
    const idc = String(c[K.CLI_ID] || "").trim();
    if (!idc) return;
    cliMap[idc] = {
      nome: String(c[K.CLI_NOME] || "").trim(),
      tel: String(c[K.CLI_TEL] || "").trim(),
      email: String(c[K.CLI_EMAIL] || "").trim()
    };
  });

  const avsEnriched = avs.map(a => {
    const idc = String(a[K.AVAL_ID_CLIENTE] || "").trim();
    const info = cliMap[idc] || null;
    return {
      ...a,
      Cliente_Nome: String(info ? info.nome : (idc || "")).replace(/^cliente\s+/i, "").trim()
    };
  });

  const nota_media = _pdf_calcNotaMedia_(avsEnriched);
  const clientes_nomes = Array.from(new Set(
    avsEnriched.map(a => String(a.Cliente_Nome || "").trim()).filter(Boolean)
  ));

  const bg_data_url = _pdf_getBackgroundDataUrl_();

  return {
    id_visita: idv,
    id_imovel: idImovel,
    data_visita: String(fato[K.FATO_DATA] || "").trim(),
    id_agendamento: idAg || "",
    anexo_ficha_visita: String(fato[K.FATO_ANEXO] || "").trim(),

    nota_media: nota_media,

    fato: fato,
    imovel: imovel,
    agenda: agenda,
    avaliacoes: avsEnriched,
    clientes_nomes,
    bg_data_url
  };
}

/* =========================
   Geração do PDF (template)
========================= */
function PDF_generatePdfVisita_v1(idVisita) {
  const payload = PDF_getVisitaPayload_v1(idVisita);
  if (!payload) throw new Error("Parâmetro visitaId ausente ou visita não encontrada.");

  const tpl = HtmlService.createTemplateFromFile("Pdf_Visita_Template");
  tpl.data = payload;
  const html = tpl.evaluate().getContent();

  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF);
  const name = `Visita_${payload.id_visita}.pdf`;
  blob.setName(name);

  const folder = DriveApp.getFolderById(PDF_getFolderId_());

  // remove anterior com mesmo nome (evita duplicação)
  const existing = folder.getFilesByName(name);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  return { ok: true, fileId: file.getId(), name: file.getName(), url: file.getUrl() };
}

/* =========================
   Ponte SAFE chamada pela Web
========================= */
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

    const res = PDF_generatePdfVisita_v1(visitaId);
    return { ok: true, ...res };

  } catch (e) {
    return { ok: false, message: e && e.message ? e.message : String(e) };
  }
}
