/**
 * FollowUpService
 * Rotina de follow-ups: e-mail diário + sincronização com Google Agenda.
 */

var FU_CFG = {
  EMAIL: 'gustavohenriquer598@gmail.com',
  CALENDAR_ID: 'gustavohenriquer598@gmail.com',
  EVENT_KEY_PREFIX: 'CRM_FOLLOWUP',
  DEFAULT_HOUR_START: 8,
  DAILY_HANDLER: 'FU_dailyFollowUpJob_',
  SHEETS: [
    { name: 'Leads_Compradores', idCandidates: ['ID', 'Id', 'id'], title: 'Lead Comprador' },
    { name: 'Leads_Vendedores', idCandidates: ['ID', 'Id', 'id'], title: 'Lead Vendedor' },
    { name: 'Fato_Visitas', idCandidates: ['Id_Visita', 'ID_Visita', 'id_visita'], title: 'Visita' },
    { name: 'Fato_Proposta', idCandidates: ['Id_Proposta', 'ID_ProPOSTA', 'id_proposta'], title: 'Proposta' },
    { name: 'Fato_Venda', idCandidates: ['Id_Venda', 'id_venda'], title: 'Venda' },
    { name: 'Fato_Captacao', idCandidates: ['Código', 'Codigo', 'ID', 'Id', 'id'], title: 'Captação' }
  ],
  FOLLOWUP_CANDIDATES: ['Próximo Follow-up', 'Proximo Follow-up', 'Próxima Data de Contato', 'Proxima Data de Contato', 'Follow-up', 'Follow up'],
  DATE_CANDIDATES: ['Data', 'DataCadastro', 'Data_Visita', 'Data de Entrada', 'Data Entrada'],
  VALUE_CANDIDATES: ['Valor', 'Valor da Proposta', 'Valor da Venda', 'Preço', 'Preco']
};

function FU_installDailyTrigger_6am() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === FU_CFG.DAILY_HANDLER) ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger(FU_CFG.DAILY_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  return { ok: true, handler: FU_CFG.DAILY_HANDLER, atHour: 6 };
}

function FU_dailyFollowUpJob_() {
  ensureSchema_();
  var itemsToday = FU_collectTodayFollowUpsDetailed_();
  FU_sendDailyEmail_(itemsToday);
  var sync = FU_syncAllFutureFollowUpsToCalendar_();
  return { ok: true, todayCount: itemsToday.length, calendar: sync };
}

function FU_syncFollowUpForRecord_(sheetName, idCol, idVal, obj) {
  try {
    if (!FU_isMonitoredSheet_(sheetName)) return { ok: true, skipped: true };

    var record = obj || DataService.getById(sheetName, idCol, idVal);
    if (!record) return { ok: true, skipped: true, reason: 'record_not_found' };

    var item = FU_buildFollowUpItem_(sheetName, record, idCol, idVal);
    if (!item) return { ok: true, skipped: true, reason: 'no_followup_date' };

    return FU_upsertCalendarEventForItem_(item);
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

function FU_syncAllFutureFollowUpsToCalendar_() {
  var today = FU_startOfDay_(new Date());
  var items = FU_collectFollowUpsByRange_(today, null);
  var created = 0, updated = 0, failed = 0;

  items.forEach(function (item) {
    var r = FU_upsertCalendarEventForItem_(item);
    if (!r || !r.ok) { failed++; return; }
    if (r.action === 'created') created++; else updated++;
  });

  return { ok: true, total: items.length, created: created, updated: updated, failed: failed };
}

function FU_collectTodayFollowUpsDetailed_() {
  var today = FU_startOfDay_(new Date());
  var end = FU_endOfDay_(today);
  return FU_collectFollowUpsByRange_(today, end);
}

function FU_sendDailyEmail_(items) {
  var tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  var dateLabel = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  var subject = 'CRM • Follow-ups do dia (' + dateLabel + ') • ' + items.length + ' item(ns)';

  var html = [];
  html.push('<h2>Follow-ups do dia - ' + dateLabel + '</h2>');

  if (!items.length) {
    html.push('<p>Nenhum follow-up para hoje.</p>');
  } else {
    var grouped = {};
    items.forEach(function (it) {
      if (!grouped[it.sheetName]) grouped[it.sheetName] = [];
      grouped[it.sheetName].push(it);
    });

    Object.keys(grouped).forEach(function (sheetName) {
      html.push('<h3>' + sheetName + ' (' + grouped[sheetName].length + ')</h3>');
      html.push('<ul>');
      grouped[sheetName].forEach(function (it) {
        html.push('<li>' + FU_escapeHtml_(FU_buildHumanLine_(it)) + '</li>');
      });
      html.push('</ul>');
    });
  }

  MailApp.sendEmail({
    to: FU_CFG.EMAIL,
    subject: subject,
    htmlBody: html.join('')
  });

  return { ok: true, sent: items.length };
}

function FU_collectFollowUpsByRange_(start, end) {
  var out = [];
  FU_CFG.SHEETS.forEach(function (cfg) {
    var rows = FU_readSheetObjects_(cfg.name);
    rows.forEach(function (row) {
      var item = FU_buildFollowUpItem_(cfg.name, row, null, null);
      if (!item || !item.followUpDate) return;
      if (start && item.followUpDate < FU_startOfDay_(start)) return;
      if (end && item.followUpDate > FU_endOfDay_(end)) return;
      out.push(item);
    });
  });
  out.sort(function (a, b) { return a.followUpDate - b.followUpDate; });
  return out;
}

function FU_buildFollowUpItem_(sheetName, row, idCol, idVal) {
  if (!row) return null;

  var followUpRaw = FU_pick_(row, FU_CFG.FOLLOWUP_CANDIDATES);
  var followUpDate = FU_parseDateAny_(followUpRaw);
  if (!followUpDate) return null;

  var recordId = String(idVal || '').trim();
  if (!recordId) {
    var cands = idCol ? [idCol] : FU_getSheetConfig_(sheetName).idCandidates;
    recordId = FU_pick_(row, cands);
  }
  recordId = String(recordId || '').trim();
  if (!recordId) return null;

  var details = FU_buildHierarchyDetails_(sheetName, row);
  var key = FU_buildEventKey_(sheetName, recordId);

  return {
    key: key,
    sheetName: sheetName,
    recordId: recordId,
    followUpDate: FU_startOfDay_(followUpDate),
    followUpRaw: followUpRaw,
    row: row,
    details: details
  };
}

function FU_buildHierarchyDetails_(sheetName, row) {
  var base = {
    data: FU_pick_(row, FU_CFG.DATE_CANDIDATES),
    valor: FU_pick_(row, FU_CFG.VALUE_CANDIDATES),
    status: FU_pick_(row, ['Status', 'status', 'Situação', 'Situacao']),
    cliente: FU_pick_(row, ['Nome Completo', 'Nome', 'Cliente', 'Proprietario']),
    telefone: FU_pick_(row, ['Telefone', 'Celular']),
    bairro: FU_pick_(row, ['Bairro']),
    endereco: FU_pick_(row, ['Endereço', 'Endereco'])
  };

  try {
    if (sheetName === 'Fato_Proposta') {
      var idVisita = FU_pick_(row, ['Id_Visita', 'id_visita']);
      var visita = idVisita && typeof PV_getVisitaContextById_v1 === 'function' ? PV_getVisitaContextById_v1(idVisita) : null;
      return {
        tipo: 'proposta',
        idVisita: idVisita,
        modalidade: FU_pick_(row, ['Modalidade de Pagamento']),
        visita: FU_slimVisita_(visita),
        base: base
      };
    }

    if (sheetName === 'Fato_Venda') {
      var idProposta = FU_pick_(row, ['Id_Proposta', 'id_proposta']);
      var propostaCtx = idProposta && typeof PV_getPropostaContextById_v1 === 'function' ? PV_getPropostaContextById_v1(idProposta) : null;
      return {
        tipo: 'venda',
        idProposta: idProposta,
        formaPagamento: FU_pick_(row, ['Forma de Pagamento']),
        comissao: FU_pick_(row, ['Comissão', 'Comissao']),
        proposta: propostaCtx ? propostaCtx.proposta : null,
        visita: propostaCtx ? FU_slimVisita_(propostaCtx.visita) : null,
        base: base
      };
    }

    if (sheetName === 'Fato_Visitas') {
      var idVisita2 = FU_pick_(row, ['Id_Visita', 'id_visita']);
      var visita2 = idVisita2 && typeof PV_getVisitaContextById_v1 === 'function' ? PV_getVisitaContextById_v1(idVisita2) : null;
      return { tipo: 'visita', visita: FU_slimVisita_(visita2), base: base };
    }

    if (sheetName === 'Fato_Captacao') {
      return {
        tipo: 'captacao',
        tipoImovel: FU_pick_(row, ['Tipo']),
        captadores: FU_pick_(row, ['Captadores']),
        base: base
      };
    }

    return { tipo: 'lead', base: base };
  } catch (e) {
    return { tipo: 'generico', base: base, erroContexto: String(e && e.message || e) };
  }
}

function FU_slimVisita_(visita) {
  if (!visita) return null;
  return {
    id_visita: visita.id_visita || '',
    id_agendamento: visita.id_agendamento || '',
    id_imovel: visita.id_imovel || '',
    data_visita: visita.data_visita || '',
    tipo_visita: visita.tipo_visita || '',
    nota_media: visita.nota_media || '',
    clientes_nomes: visita.clientes_nomes || [],
    imovel: visita.imovel || null,
    agenda: visita.agenda || null
  };
}

function FU_upsertCalendarEventForItem_(item) {
  var cal = FU_getCalendar_();
  var dayStart = FU_startOfDay_(item.followUpDate);
  var dayEnd = FU_endOfDay_(item.followUpDate);
  var events = cal.getEvents(dayStart, dayEnd);
  var match = null;

  for (var i = 0; i < events.length; i++) {
    var d = events[i].getDescription() || '';
    if (d.indexOf(item.key) !== -1) { match = events[i]; break; }
  }

  var title = FU_buildEventTitle_(item);
  var desc = FU_buildEventDescription_(item);

  if (match) {
    match.setTitle(title);
    match.setDescription(desc);
    return { ok: true, action: 'updated', id: match.getId() };
  }

  var ev = cal.createAllDayEvent(title, dayStart, { description: desc });
  return { ok: true, action: 'created', id: ev.getId() };
}

function FU_buildEventTitle_(item) {
  var cfg = FU_getSheetConfig_(item.sheetName);
  return 'Follow-up ' + (cfg.title || item.sheetName) + ' #' + item.recordId;
}

function FU_buildEventDescription_(item) {
  var lines = [];
  lines.push(item.key);
  lines.push('Origem: ' + item.sheetName);
  lines.push('ID Registro: ' + item.recordId);
  lines.push('Follow-up: ' + FU_fmtDate_(item.followUpDate));
  lines.push('Resumo: ' + FU_buildHumanLine_(item));

  if (item.details && item.details.base) {
    var b = item.details.base;
    if (b.cliente) lines.push('Cliente: ' + b.cliente);
    if (b.telefone) lines.push('Telefone: ' + b.telefone);
    if (b.valor) lines.push('Valor: ' + b.valor);
    if (b.status) lines.push('Status: ' + b.status);
    if (b.bairro) lines.push('Bairro: ' + b.bairro);
  }

  return lines.join('\n');
}

function FU_buildHumanLine_(item) {
  var bits = [];
  bits.push(item.sheetName + ' #' + item.recordId);
  if (item.details && item.details.base) {
    var b = item.details.base;
    if (b.cliente) bits.push('Cliente: ' + b.cliente);
    if (b.telefone) bits.push('Tel: ' + b.telefone);
    if (b.status) bits.push('Status: ' + b.status);
    if (b.valor) bits.push('Valor: ' + b.valor);
  }

  if (item.details && item.details.tipo === 'proposta' && item.details.idVisita) {
    bits.push('Visita: ' + item.details.idVisita);
  }
  if (item.details && item.details.tipo === 'venda' && item.details.idProposta) {
    bits.push('Proposta: ' + item.details.idProposta);
  }

  return bits.join(' • ');
}

function FU_getCalendar_() {
  return CalendarApp.getCalendarById(FU_CFG.CALENDAR_ID) || CalendarApp.getDefaultCalendar();
}

function FU_getSheetConfig_(sheetName) {
  for (var i = 0; i < FU_CFG.SHEETS.length; i++) {
    if (FU_CFG.SHEETS[i].name === sheetName) return FU_CFG.SHEETS[i];
  }
  return { name: sheetName, idCandidates: ['ID', 'Id', 'id'], title: sheetName };
}

function FU_isMonitoredSheet_(sheetName) {
  return FU_CFG.SHEETS.some(function (s) { return s.name === sheetName; });
}

function FU_readSheetObjects_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  var headers = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function (h) { return String(h || '').trim(); });
  var values = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();

  return values.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { if (h) obj[h] = row[i]; });
    return obj;
  });
}

function FU_pick_(obj, candidates) {
  if (!obj) return '';
  var keys = {};
  Object.keys(obj).forEach(function (k) { keys[FU_norm_(k)] = k; });

  for (var i = 0; i < (candidates || []).length; i++) {
    var want = FU_norm_(candidates[i]);
    if (keys.hasOwnProperty(want)) return String(obj[keys[want]] || '').trim();
  }
  return '';
}

function FU_norm_(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function FU_parseDateAny_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  var s = String(v || '').trim();
  if (!s) return null;

  var d = new Date(s);
  if (!isNaN(d)) return d;

  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));

  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  return null;
}

function FU_startOfDay_(d) {
  var x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function FU_endOfDay_(d) {
  var x = new Date(d);
  x.setHours(23, 59, 59, 999);
  return x;
}

function FU_fmtDate_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/Sao_Paulo', 'dd/MM/yyyy');
}

function FU_buildEventKey_(sheetName, recordId) {
  return FU_CFG.EVENT_KEY_PREFIX + '|' + sheetName + '|' + recordId;
}

function FU_escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
