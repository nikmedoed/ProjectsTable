function switchReport(forceReport) {
  let rep = getReportState()
  if (forceReport === rep) {
    return
  }
  let showReport = !rep

  let blocks = getBlockSheets()
  blocks.forEach(s => hideRows(s, showReport))

  let sname = SSheet.getName()
  let emojiRegex = new RegExp(REPORT_EMOJI, "g");
  sname = sname.replace(emojiRegex, "").trim();
  if (showReport) {
    sname = `${REPORT_EMOJI} ${sname}`;
  }
  SSheet.rename(sname)

  storeReportState(showReport)
}


function hideRows(sheet, hide = false) {
  var data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (var i = ROW_START - 1; i < data.length; i++) {
    if (hide) {
      if (!data[i][0]) {
        sheet.hideRow(sheet.getRange(i + 1, 1));
      }
    } else {
      sheet.unhideRow(sheet.getRange(i + 1, 1));
    }
  }
}
