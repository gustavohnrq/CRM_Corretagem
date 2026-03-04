var UiService = (function () {
  function include(filename) {
    try {
      return HtmlService.createHtmlOutputFromFile(String(filename || '').trim()).getContent();
    } catch (err) {
      Logger.log('[UiService.include] Falha ao incluir arquivo "' + filename + '": ' + (err && err.message ? err.message : err));
      return '';
    }
  }
  return { include: include };
})();
