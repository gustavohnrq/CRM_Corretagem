var UiService = (function () {
  function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }
  return { include: include };
})();