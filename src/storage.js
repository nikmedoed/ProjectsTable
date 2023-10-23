function isValidDate(d) {
  var date = new Date(d);
  return date instanceof Date && !isNaN(date);
}


TIMELINE_END_DATE = "TIMELINE_END_DATE"
TIMELINE_START_DATE = "TIMELINE_START_DATE"

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
    PropertiesService.getScriptProperties().setProperty(key, date);
  } else {
    throw new Error('Некорректная дата');
  }
  return date
}

function getDate(key) {
  var storedDate = PropertiesService.getScriptProperties().getProperty(key);
  return storedDate ? new Date(storedDate) : "";
}
