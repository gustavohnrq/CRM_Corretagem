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
    { name: 'Base_Clientes', idCandidates: ['ID', 'Id', 'id'], title: 'Cliente' },
    { name: 'Leads_Compradores', idCandidates: ['ID', 'Id', 'id'], title: 'Lead Comprador' },
    { name: 'Leads_Vendedores', idCandidates: ['ID', 'Id', 'id'], title: 'Lead Vendedor' },
    { name: 'Agenda_Visitas', idCandidates: ['ID', 'Id', 'id'], title: 'Visita' },
    { name: 'Fato_Visitas', idCandidates: ['Id_Visita', 'ID_Visita', 'id_visita'], title: 'Visita' },
    { name: 'Fato_Proposta', idCandidates: ['Id_Proposta', 'ID_ProPOSTA', 'id_proposta'], title: 'Proposta' },
    { name: 'Fato_Venda', idCandidates: ['Id_Venda', 'id_venda'], title: 'Venda' },
    { name: 'Fato_Captacao', idCandidates: ['Código', 'Codigo', 'ID', 'Id', 'id'], title: 'Captação' }
  ],
  FOLLOWUP_CANDIDATES: ['Próximo Follow-up'],
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
  Logger.log('[FollowUp] Iniciando rotina diária. Itens do dia: ' + itemsToday.length);
  FU_sendDailyEmail_(itemsToday);
  var sync = FU_syncAllFutureFollowUpsToCalendar_();
  Logger.log('[FollowUp] Rotina diária concluída com sucesso. Total=' + sync.total + ' | Criados=' + sync.created + ' | Atualizados=' + sync.updated + ' | Falhas=' + sync.failed);
  return { ok: true, todayCount: itemsToday.length, calendar: sync };
}


function FU_syncAgendaNow_v1() {
  ensureSchema_();
  Logger.log('[FollowUp] Sincronização manual da agenda solicitada via menu.');
  var sync = FU_syncAllFutureFollowUpsToCalendar_();
  Logger.log('[FollowUp] Sincronização manual concluída com sucesso. Total=' + sync.total + ' | Criados=' + sync.created + ' | Atualizados=' + sync.updated + ' | Falhas=' + sync.failed);
  return { ok: true, source: 'manual_menu', calendar: sync };
}
function FU_syncFollowUpForRecord_(sheetName, idCol, idVal, obj) {
  try {
    if (!FU_isMonitoredSheet_(sheetName)) return { ok: true, skipped: true };
    if (!FU_sheetHasFollowDateColumn_(sheetName)) return { ok: true, skipped: true, reason: 'sheet_without_follow_date' };

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
  var bySheet = {};

  items.forEach(function (item) {
    if (!bySheet[item.sheetName]) bySheet[item.sheetName] = { total: 0, created: 0, updated: 0, failed: 0 };
    bySheet[item.sheetName].total++;

    var r = FU_upsertCalendarEventForItem_(item);
    if (!r || !r.ok) {
      failed++;
      bySheet[item.sheetName].failed++;
      Logger.log('[FollowUp] Falha ao sincronizar calendário: ' + FU_buildHumanLine_(item));
      return;
    }

    if (r.action === 'created') {
      created++;
      bySheet[item.sheetName].created++;
    } else {
      updated++;
      bySheet[item.sheetName].updated++;
    }
  });

  Logger.log('[FollowUp] Resumo por aba: ' + JSON.stringify(bySheet));
  return { ok: true, total: items.length, created: created, updated: updated, failed: failed, bySheet: bySheet };
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

  var preview = items.slice(0, 20).map(function (it) { return FU_buildHumanLine_(it); }).join(' | ');
  Logger.log('[FollowUp] Enviando e-mail para: ' + FU_CFG.EMAIL + ' | assunto: ' + subject + ' | qtd: ' + items.length);
  if (preview) Logger.log('[FollowUp] Itens (preview): ' + preview);

  MailApp.sendEmail({
    to: FU_CFG.EMAIL,
    subject: subject,
    htmlBody: html.join('')
  });

  return { ok: true, sent: items.length };
}


function FU_sheetHasFollowDateColumn_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return false;
  var lc = sh.getLastColumn();
  if (lc < 1) return false;
  var headers = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(function (h) { return String(h || '').trim(); });
  if (sheetName === 'Agenda_Visitas') return headers.indexOf('Data') !== -1;
  return headers.indexOf('Próximo Follow-up') !== -1;
}

function FU_collectFollowUpsByRange_(start, end) {
  var out = [];
  FU_CFG.SHEETS.forEach(function (cfg) {
    if (!FU_sheetHasFollowDateColumn_(cfg.name)) return;
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

  var follow = FU_getFollowDateForRow_(sheetName, row);
  var followUpRaw = follow.raw;
  var followUpDate = follow.date;
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


function FU_getFollowDateForRow_(sheetName, row) {
  if (!row) return { raw: '', date: null };

  if (sheetName === 'Agenda_Visitas') {
    var rawAgenda = String(row['Data'] || '').trim();
    return { raw: rawAgenda, date: FU_parseDateAny_(rawAgenda) };
  }

  var raw = String(row['Próximo Follow-up'] || '').trim();
  return { raw: raw, date: FU_parseDateAny_(raw) };
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

    if (sheetName === 'Base_Clientes') {
      return {
        tipo: 'cliente',
        origem: FU_pick_(row, ['Origem']),
        statusAtual: FU_pick_(row, ['Status Atual']),
        perfil: FU_pick_(row, ['Perfil (Moradia/Investimento)']),
        regiaoInteresse: FU_pick_(row, ['Região de Interesse']),
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
  var window = FU_resolveTimeWindow_(item);
  var startAt = window.startAt;
  var endAt = window.endAt;
  var windowLabel = window.windowLabel;

  var title = FU_buildEventTitle_(item);
  var desc = FU_buildEventDescription_(item) + '\nJanela: ' + windowLabel;
  var calendar = FU_getCalendarOrThrow_();

  Logger.log('[FollowUp] Payload evento => calendar=' + FU_CFG.CALENDAR_ID + ' | key=' + item.key + ' | title=' + title + ' | date=' + FU_fmtDate_(item.followUpDate) + ' | start=' + FU_fmtHour_(startAt) + ' | end=' + FU_fmtHour_(endAt));

  var matches = FU_findCalendarEventsByKey_(calendar, item.key);
  var match = matches.length ? matches[0] : null;

  if (match) {
    match.setTitle(title);
    match.setDescription(desc);
    match.setTime(startAt, endAt);

    if (matches.length > 1) {
      for (var i = 1; i < matches.length; i++) {
        try {
          matches[i].deleteEvent();
          Logger.log('[FollowUp] Evento duplicado removido: key=' + item.key + ' | id=' + matches[i].getId());
        } catch (e) {}
      }
    }

    FU_setEventIdByKey_(item.key, match.getId());
    Logger.log('[FollowUp] Evento atualizado: ' + title + ' | key: ' + item.key + ' | eventId=' + match.getId());
    return { ok: true, action: 'updated', id: match.getId() };
  }

  var created = calendar.createEvent(title, startAt, endAt, { description: desc });
  FU_setEventIdByKey_(item.key, created.getId());
  Logger.log('[FollowUp] Evento criado: ' + title + ' | key: ' + item.key + ' | eventId=' + created.getId());
  return { ok: true, action: 'created', id: created.getId() };
}

function FU_resolveTimeWindow_(item) {
  var defaultStart = FU_withTime_(item.followUpDate, 6, 0, 0);
  var defaultEnd = FU_withTime_(item.followUpDate, 18, 0, 0);
  var rawHour = FU_pick_(item.row, ['Hora']);
  var parsedHour = FU_parseTimeAny_(rawHour);

  if (item.sheetName === 'Agenda_Visitas' && parsedHour) {
    var start = FU_withTime_(item.followUpDate, parsedHour.hh, parsedHour.mm, 0);
    var end = FU_withTime_(item.followUpDate, parsedHour.hh + 1, parsedHour.mm, 0);
    return {
      startAt: start,
      endAt: end,
      hourLabel: parsedHour.label,
      windowLabel: parsedHour.label + ' às ' + FU_fmtHour_(end)
    };
  }

  return {
    startAt: defaultStart,
    endAt: defaultEnd,
    hourLabel: parsedHour ? parsedHour.label : '-',
    windowLabel: '06:00 às 18:00'
  };
}


function FU_buildEventTitle_(item) {
  if (item.sheetName === 'Agenda_Visitas') {
    var tipoVisita = FU_pick_(item.row, ['Venda ou Cap?']) || 'Visita';
    return 'Visita de ' + tipoVisita;
  }

  var cfg = FU_getSheetConfig_(item.sheetName);
  var base = (cfg.title || item.sheetName || 'Registro');
  return 'Follow Up ' + base + ' - ' + item.recordId;
}


function FU_getCalendarOrThrow_() {
  var cal = CalendarApp.getCalendarById(FU_CFG.CALENDAR_ID);
  if (!cal) throw new Error('Calendário não encontrado para CALENDAR_ID=' + FU_CFG.CALENDAR_ID);
  return cal;
}

function FU_findCalendarEventsByKey_(calendar, key) {
  var mapped = FU_getMappedEventByKey_(calendar, key);
  var out = mapped ? [mapped] : [];

  var now = FU_startOfDay_(new Date());
  var from = FU_withTime_(new Date(now.getFullYear() - 1, 0, 1), 0, 0, 0);
  var to = FU_withTime_(new Date(now.getFullYear() + 5, 11, 31), 23, 59, 59);
  var searched = calendar.getEvents(from, to, { search: key }) || [];

  for (var i = 0; i < searched.length; i++) {
    var ev = searched[i];
    var desc = String(ev.getDescription() || '');
    if (desc.indexOf(key) === -1) continue;
    if (!out.some(function (x) { return x.getId() === ev.getId(); })) out.push(ev);
  }

  if (out.length) FU_setEventIdByKey_(key, out[0].getId());
  return out;
}

function FU_getMappedEventByKey_(calendar, key) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(FU_eventPropKey_(key));
  if (!id) return null;

  try {
    var ev = calendar.getEventById(id);
    if (!ev) return null;
    var desc = String(ev.getDescription() || '');
    if (desc.indexOf(key) === -1) return null;
    return ev;
  } catch (e) {
    return null;
  }
}

function FU_setEventIdByKey_(key, id) {
  if (!id) return;
  PropertiesService.getScriptProperties().setProperty(FU_eventPropKey_(key), id);
}

function FU_eventPropKey_(key) {
  return 'FU_EVENT_ID|' + key;
}

function FU_parseTimeAny_(v) {
  var s = String(v || '').trim();
  if (!s) return null;

  var m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!m) m = s.match(/^(\d{1,2})h(\d{2})$/i);
  if (!m) return null;

  var hh = Math.max(0, Math.min(23, Number(m[1])));
  var mm = Math.max(0, Math.min(59, Number(m[2])));
  return { hh: hh, mm: mm, label: (hh < 10 ? '0' : '') + hh + ':' + (mm < 10 ? '0' : '') + mm };
}

function FU_fmtHour_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/Sao_Paulo', 'HH:mm');
}


function FU_withTime_(d, hh, mm, ss) {
  var x = new Date(d);
  x.setHours(hh || 0, mm || 0, ss || 0, 0);
  return x;
}

function FU_buildEventDescription_(item) {
  var lines = [];
  lines.push(item.key);
  lines.push('Origem: ' + item.sheetName);
  lines.push('ID Registro: ' + item.recordId);
  lines.push('Follow-up: ' + FU_fmtDate_(item.followUpDate));
  var hora = FU_pick_(item.row, ['Hora']);
  if (hora) lines.push('Hora: ' + hora);
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
  return FU_getCalendarOrThrow_();
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
  if (v instanceof Date && !isNaN(v)) return FU_startOfDay_(v);
  var s = String(v || '').trim();
  if (!s) return null;

  // Prioridade BR: DD/MM/AAAA
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return FU_startOfDay_(new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])));

  // ISO: AAAA-MM-DD
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return FU_startOfDay_(new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])));

  var d = new Date(s);
  if (!isNaN(d)) return FU_startOfDay_(d);

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
