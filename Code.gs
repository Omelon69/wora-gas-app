function doGet(e) {
  const page = (e && e.parameter.page) ? e.parameter.page : 'index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle('WoraCRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
