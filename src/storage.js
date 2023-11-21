function isValidDate(d) {
  var date = new Date(d);
  return date instanceof Date && !isNaN(date);
}

const TIMELINE_END_DATE = "TIMELINE_END_DATE"
const TIMELINE_START_DATE = "TIMELINE_START_DATE"
const REPORT_STATE = "REPORT_STATE"
const PRESENTATION_ID = "PRESENTATION_ID";


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


function storePresentationIdOrLink(input) {
  let id = extractId(input);
  if (id) {
    STORAGE().setProperty(PRESENTATION_ID, id);
  } else {
    throw new Error('Некорректный ID или ссылка на презентацию');
  }
}

function extractId(input) {
  if (input.includes('https://')) {
    let match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
  }
  return input;
}

function getPresentationId() {
  return STORAGE().getProperty(PRESENTATION_ID);
}

function getPresentationTemplateLink() {
  const id = getPresentationId();
  if (id) {
    return `https://docs.google.com/presentation/d/${id}/edit`;
  } else {
    return 'Шаблон отчёта не привязан в таблице, создайте и привяжите новый вручную.';
  }
}
