var UNBService = (function () {
  var SHEET = 'Agenda_UNB';
  var ID_COL = 'ID';

  function listAtividades(filters) {
    ensureSchema_();
    var rows = readSheetObjects_(SHEET);
    var f = filters || {};
    var status = String(f.status || '').trim().toLowerCase();

    if (status) {
      rows = rows.filter(function (r) {
        return String(r.Status || '').trim().toLowerCase() === status;
      });
    }

    rows.sort(function (a, b) {
      var da = dsParseDateAny_(a['Próximo Follow-up'] || a.Data);
      var db = dsParseDateAny_(b['Próximo Follow-up'] || b.Data);
      if (!da && !db) return Number(b[ID_COL] || 0) - Number(a[ID_COL] || 0);
      if (!da) return 1;
      if (!db) return -1;
      return da - db;
    });

    return rows;
  }

  function getById(id) {
    return DataService.getById(SHEET, ID_COL, id);
  }

  function upsertAtividade(obj) {
    ensureSchema_();
    var payload = normalizePayload_(obj || {});
    var res = DataService.upsertById(SHEET, ID_COL, payload);

    var idVal = String((res && res.id) || payload[ID_COL] || '').trim();
    if (idVal) {
      FU_syncFollowUpForRecord_(SHEET, ID_COL, idVal, payload);
    }

    return res;
  }

  function deleteAtividade(id) {
    return DataService.deleteById(SHEET, ID_COL, id);
  }

  function getDashboardData() {
    ensureSchema_();
    var all = listAtividades({});
    var today = dsStartOfDay_(new Date());
    var tomorrow = addDays_(today, 1);
    var plus7 = addDays_(today, 7);

    var abertas = all.filter(function (r) {
      return dsNorm_(r.Status || 'Aberto') !== 'concluído' && dsNorm_(r.Status || 'Aberto') !== 'concluido';
    });

    var hoje = abertas.filter(function (r) {
      var d = dsParseDateAny_(r['Próximo Follow-up'] || r.Data);
      return d && d >= today && d < tomorrow;
    });

    var semana = abertas.filter(function (r) {
      var d = dsParseDateAny_(r['Próximo Follow-up'] || r.Data);
      return d && d >= today && d < plus7;
    });

    var porMateria = {};
    abertas.forEach(function (r) {
      var m = String(r['Matéria'] || 'Sem matéria').trim() || 'Sem matéria';
      porMateria[m] = (porMateria[m] || 0) + 1;
    });

    var materias = Object.keys(porMateria)
      .sort(function (a, b) { return porMateria[b] - porMateria[a]; })
      .map(function (k) { return { materia: k, qtd: porMateria[k] }; });

    return {
      totais: {
        total: all.length,
        abertas: abertas.length,
        hoje: hoje.length,
        semana: semana.length
      },
      all: all,
      hoje: hoje,
      semana: semana,
      materias: materias
    };
  }

  function getFieldOptions() {
    ensureSchema_();
    var rows = readSheetObjects_(SHEET);
    return {
      materias: unique_(rows, 'Matéria'),
      atividades: unique_(rows, 'Atividade'),
      descricoes: unique_(rows, 'Descrição'),
      prioridades: unique_(rows, 'Prioridade'),
      status: unique_(rows, 'Status')
    };
  }

  function normalizePayload_(obj) {
    var o = Object.assign({}, obj);
    if (!o.DataCadastro) o.DataCadastro = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Sao_Paulo', 'yyyy-MM-dd');
    if (!o.Status) o.Status = 'Aberto';
    if (!o['Próximo Follow-up']) o['Próximo Follow-up'] = o.Data || '';
    return o;
  }

  function unique_(rows, key) {
    var set = {};
    (rows || []).forEach(function (r) {
      var v = String(r[key] || '').trim();
      if (v) set[v] = true;
    });
    return Object.keys(set).sort(function (a, b) { return a.localeCompare(b); });
  }

  return {
    listAtividades: listAtividades,
    getById: getById,
    upsertAtividade: upsertAtividade,
    deleteAtividade: deleteAtividade,
    getDashboardData: getDashboardData,
    getFieldOptions: getFieldOptions
  };
})();
