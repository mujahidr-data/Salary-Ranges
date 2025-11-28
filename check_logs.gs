function showRecentLogs() {
  const logs = Logger.getLog();
  console.log(logs);
  SpreadsheetApp.getUi().alert('Logs', logs.substring(0, 5000), SpreadsheetApp.getUi().ButtonSet.OK);
}
