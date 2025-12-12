// function preloadHolidayCache() {
//   const year = new Date().getFullYear();
//   const url = `https://www.hebcal.com/hebcal?v=1&cfg=json&maj=on&year=${year}&c=on&geo=IL`;
//   const response = UrlFetchApp.fetch(url);
//   const data = JSON.parse(response.getContentText());
//   const holidays = data.items
//     .filter(item => item.category === "holiday" || item.category === "festival")
//     .map(item => ({
//       date: item.date,
//       title: item.title
//     }));
//   CacheService.getScriptCache().put("holidays_" + year, JSON.stringify(holidays), 21600);
// }

// function getHolidayInfo(date) {
//   if (!date) return "";
//   if (!(date instanceof Date)) date = new Date(date);

//   const year = date.getFullYear();
//   const cached = CacheService.getScriptCache().get("holidays_" + year);
//   if (!cached) return "";

//   const holidays = JSON.parse(cached);
//   const targetDate = Utilities.formatDate(new Date(date), "GMT+3", "yyyy-MM-dd");

//   for (let i = 0; i < holidays.length; i++) {
//     if (holidays[i].date === targetDate) {
//       return holidays[i].title;
//     }
//   }
//   return "";
// }

// function isIsraelHoliday(date) {
//   return getHolidayInfo(date) !== "";
// }
