// function addBirthdays_MeidaKlalit_WithNamePrompt() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("מידע כללי");
//   if (!sheet) return;

//   const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//   const emailCol = headers.indexOf("מייל אישי") + 1;
//   const birthCol = headers.indexOf("תאריך לידה") + 1;
//   const nameCol = headers.indexOf("שם") + 1;

//   if ([emailCol, birthCol, nameCol].includes(0)) {
//     Logger.log("עמודות חסרות: ודא שקיימות 'שם', 'מייל אישי' ו'תאריך לידה'");
//     return;
//   }

//   let statusCol = headers.indexOf("סטטוס יום הולדת") + 1;
//   if (statusCol === 0) {
//     statusCol = headers.length + 1;
//     sheet.getRange(1, statusCol).setValue("סטטוס יום הולדת");
//   }

//   const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

//   // שלב 1: מצא שורות עם שם חסר או קצר מדי (פחות מ-3 תווים)
//   const missingNameIndices = [];
//   const missingNameEmails = [];

//   data.forEach((row, i) => {
//     const name = (row[nameCol - 1] || "").toString().trim();
//     if (name.length < 3) {
//       missingNameIndices.push(i);
//       missingNameEmails.push((row[emailCol - 1] || "").toString().trim());
//     }
//   });

//   if (missingNameIndices.length > 0) {
//     const ui = SpreadsheetApp.getUi();
//     const promptResponse = ui.prompt(
//       'חסרים שמות מלאים',
//       `אנא הזן כאן את השמות המלאים המתאימים (משפחה ופרטי), כל שם בשורה חדשה, לפי הסדר הבא:\n${missingNameEmails.join('\n')}`,
//       ui.ButtonSet.OK_CANCEL
//     );

//     if (promptResponse.getSelectedButton() != ui.Button.OK) {
//       Logger.log("המשתמש ביטל את הזנת השמות");
//       return;
//     }

//     const inputText = promptResponse.getResponseText();
//     const namesInput = inputText.split(/\r?\n/).map(s => s.trim());

//     if (namesInput.length !== missingNameIndices.length) {
//       ui.alert(`מספר השמות שהוזנו (${namesInput.length}) שונה ממספר האנשים שחסר להם שם (${missingNameIndices.length}). הפעל מחדש ונסה שוב.`);
//       return;
//     }

//     // שלב 2: מלא את השמות החסרים במערך הנתונים
//     namesInput.forEach((name, idx) => {
//       data[missingNameIndices[idx]][nameCol - 1] = name;
//     });
//   }

//   // שלב 3: יצירת אירועי יום הולדת ביומן
//   const calendar = CalendarApp.getCalendarById("shlomiedria@gmail.com");
//   const currentYear = new Date().getFullYear();

//   data.forEach((row, i) => {
//     let email = (row[emailCol - 1] || "").toString().trim().replace("/com", ".com").replace("/", "@");
//     const birthDate = row[birthCol - 1];
//     const name = (row[nameCol - 1] || "").toString().trim();
//     if (!email || !birthDate || !name) return;

//     const bd = new Date(birthDate);
//     if (isNaN(bd)) return;

//     const eventDate = new Date(currentYear, bd.getMonth(), bd.getDate());
//     const endDate = new Date(currentYear, bd.getMonth(), bd.getDate() + 1);

//     const existingEvents = calendar.getEvents(eventDate, endDate);
//     const alreadyExists = existingEvents.some(event =>
//       event.getDescription().toLowerCase().includes(email.toLowerCase())
//     );

//     if (alreadyExists) {
//       sheet.getRange(i + 2, statusCol).setValue("כבר ביומן");
//       return;
//     }

//     const title = `🎂 יום הולדת - ${name}`;
//     const description = `תאריך לידה: ${bd.toLocaleDateString("he-IL")}\nאימייל: ${email}`;

//     calendar.createAllDayEvent(title, eventDate, {
//       description: description,
//       recurrence: CalendarApp.newRecurrence().addYearlyRule()
//     });

//     sheet.getRange(i + 2, statusCol).setValue("נוצר");
//   });
// }
