function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C13:C16').activate();
  spreadsheet.getRange('C13:C16').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  var banding = spreadsheet.getRange('C13:C16').getBandings()[0];
  banding.setHeaderRowColor('#f7cb4d')
  .setFirstRowColor('#ffffff')
  .setSecondRowColor('#fef8e3')
  .setFooterRowColor(null);
  banding = spreadsheet.getRange('C13:C16').getBandings()[0];
  banding.setHeaderRowColor(null)
  .setFirstRowColor('#ffffff')
  .setSecondRowColor('#fef8e3')
  .setFooterRowColor(null);
};

function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B2:C9').activate();
  spreadsheet.getRange('B2:C9').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  var banding = spreadsheet.getRange('B2:C9').getBandings()[0];
  banding.setHeaderRowColor('#d9ead3')
  .setFirstRowColor('#bf9000')
  .setSecondRowColor('#fef8e3')
  .setFooterRowColor(null);
};