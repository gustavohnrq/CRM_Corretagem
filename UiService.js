var UiService = (function () {
  function include(filename) {
    var safe = String(filename || '').trim();
    var allowed = {
      Estilos: true,
      JsBase: true
    };

    if (!allowed[safe]) {
      Logger.log('[UiService.include] Include não permitido: "' + safe + '"');
      return '';
    }

    try {
      return HtmlService.createHtmlOutputFromFile(safe).getContent();
    } catch (err) {
      Logger.log('[UiService.include] Falha ao incluir arquivo "' + safe + '": ' + (err && err.message ? err.message : err));
      return '';
    }
  }
  return { include: include };
})();
