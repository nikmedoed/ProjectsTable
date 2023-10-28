function isValidDate(d) {
  var date = new Date(d);
  return date instanceof Date && !isNaN(date);
}

const TIMELINE_END_DATE = "TIMELINE_END_DATE"
const TIMELINE_START_DATE = "TIMELINE_START_DATE"
const REPORT_STATE = "REPORT_STATE"

const STORAGE = PropertiesService.getDocumentProperties

function storeReportState(state = false) {
  STORAGE().setProperty("REPORT_STATE", state.toString());
}

function getReportState() {
  return STORAGE().getProperty("REPORT_STATE") === 'true';
}


function storeTimelineEndDate(date) {
  storeDate(TIMELINE_END_DATE, date)
}

function getTimelineEndDate() {
  return getDate(TIMELINE_END_DATE)
}


function storeTimelineStartDate(date) {
  storeDate(TIMELINE_START_DATE, date)
}

function getTimelineStartDate() {
  return getDate(TIMELINE_START_DATE)
}


function storeDate(key, date) {
  if (isValidDate(date)) {
    STORAGE().setProperty(key, date);
  } else {
    throw new Error('Некорректная дата');
  }
  return date
}

function getDate(key) {
  var storedDate = STORAGE().getProperty(key);
  return storedDate ? new Date(storedDate) : "";
}
