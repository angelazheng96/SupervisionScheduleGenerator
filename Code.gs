function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Schedule')
    .addItem('Sort teachers sheet', 'sortTeachersSheet')
    .addItem('Format teachers sheet', 'formatTeachersSheet')
    .addToUi();
}

// Angela - base function: generates main schedule
function generateSchedule(sheetName) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let schedule = sheet.getSheetByName(sheetName).getRange('A4:I23').getValues();
  let teachersArray = getTeachersArray();
  let organizedTeachers = organizedTeacherPeriods(teachersArray);
  console.log("ORGANIZED TEACHERS");
  console.log(organizedTeachers);
  let averageRatio = getAverageRatio(teachersArray);
  let teachersNeedingShifts = getTeachersNeedingShifts(organizedTeachers, teachersArray, averageRatio);

  let currentPeriod, prevPeriod, currentDay, currentDutyStr, coordinates;
  let availablePeriods, availableTeachers, preferredTeachers, unpreferredTeachers;

  // repeats it all twice, so it does all day 1s then all day 2s
  for (let i = 0; i < 2; i++) {
    currentDay = schedule[i][2];

    // cycles through each alternating row of the schedule (first all day 1s, then all day 2s)
    for (let currentDuty = i; currentDuty < schedule.length; currentDuty += 2) {

      // sets values for the current duty's period and day
      currentPeriod = schedule[currentDuty][3];
      currentDutyStr = schedule[Math.floor(currentDuty / 2) * 2][0];

      console.log("current period: " + currentPeriod);
      console.log("current day: " + currentDay);
      console.log("current duty: " + currentDutyStr);

      // if it is a new period, then it needs to update the availablePeriods and availableTeachers arrays
      if (currentPeriod != prevPeriod) {
        // figures out which prep periods are available for current duty
        availablePeriods = getAvailablePeriods(currentPeriod, currentDay);

        // filters teachers based on who is available for the current period(s)
        availableTeachers = filterbyAvailability(organizedTeachers, availablePeriods);

        // filters teachers based on who needs shifts
        teachersNeedingShifts = getTeachersNeedingShifts(availableTeachers, teachersArray, averageRatio);
      }

      console.log("AVAILABLE PERIODS");
      console.log(availablePeriods);

      console.log("AVAILABLE TEACHERS");
      console.log(availableTeachers);

      console.log("TEACHERS NEEDING SHIFTS");
      console.log(teachersNeedingShifts);

      // cycles through days of week for current shift
      for (let dayOfWeek = 0; dayOfWeek < 5; dayOfWeek++) {

        // if there is at least 1 teacher who needs more shifts
        if (countTotalElements(teachersNeedingShifts) > 0) {

          // filters teachers who need shifts based on preferred duties
          preferredTeachers = filterByPreferredDuty(teachersNeedingShifts, teachersArray, currentDutyStr);

          // ideal scenario: find an available teacher needing more shifts who prefers this duty
          if (countTotalElements(preferredTeachers) > 0) {
            coordinates = randomElement(preferredTeachers);
            currentTeacher = preferredTeachers[coordinates[0]][coordinates[1]];

          // if there are no available teachers who prefer this duty
          } else {
            unpreferredTeachers = filterByUnpreferredDuty(teachersNeedingShifts, teachersArray, currentDutyStr);

            // if there are teachers who don't dislike this duty
            if (countTotalElements(unpreferredTeachers) > 0) {
              coordinates = randomElement(unpreferredTeachers);
              currentTeacher = unpreferredTeachers[coordinates[0]][coordinates[1]];

            // otherwise, if everyone dislikes this duty, still finds someone to do it
            } else {
              coordinates = randomElement(availableTeachers);
              currentTeacher = availableTeachers[coordinates[0]][coordinates[1]];
            }
          }

        // if none of the teachers need more shifts
        } else {

          // filters teachers who don't need shifts based on preferred duties
          preferredTeachers = filterByPreferredDuty(availableTeachers, teachersArray, currentDutyStr);

          // ideal scenario: find an available teacher who prefers this duty
          if (countTotalElements(preferredTeachers) > 0) {
            coordinates = randomElement(preferredTeachers);
            currentTeacher = preferredTeachers[coordinates[0]][coordinates[1]];

          // if there are no available teachers who prefer this duty
          } else {
            unpreferredTeachers = filterByUnpreferredDuty(availableTeachers, teachersArray, currentDutyStr);

            // if there are teachers who don't dislike this duty
            if (countTotalElements(unpreferredTeachers) > 0) {
              coordinates = randomElement(unpreferredTeachers);
              currentTeacher = unpreferredTeachers[coordinates[0]][coordinates[1]];

            // otherwise, if everyone dislikes this duty, still finds someone to do it
            } else {
              coordinates = randomElement(availableTeachers);
              currentTeacher = availableTeachers[coordinates[0]][coordinates[1]];
            }
          }

        }

        console.log("PREFERRED TEACHERS");
        console.log(preferredTeachers);
        console.log("UNPREFERRED TEACHERS");
        console.log(unpreferredTeachers);
        console.log("coordinates");
        console.log(coordinates);

        // adds the teacher to the schedule
        schedule[currentDuty][dayOfWeek + 4] = currentTeacher;
        console.log('added: ' + currentTeacher);

        // removes teacher from the availableTeachers & organizedTeachers arrays
        organizedTeachers = removeElement(organizedTeachers, availablePeriods[coordinates[0]] - 1, currentTeacher);
        availableTeachers = removeElement(availableTeachers, coordinates[0], currentTeacher);
        teachersNeedingShifts = removeElement(teachersNeedingShifts, coordinates[0], currentTeacher);
      }

      prevPeriod = currentPeriod;

    }
  }

  // puts values into spreadsheet
  sheet.getSheetByName(sheetName).getRange('A4:I23').setValues(schedule);
}

// Angela - formats teachers sheet
function formatTeachersSheet() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let teacherSheet = sheet.getSheetByName('Teachers').getDataRange();

  teacherSheet.setFontColor('black');
  teacherSheet.setBackgroundColor('white');
  teacherSheet.setFontWeight('normal');
  teacherSheet.setFontStyle('normal');
  teacherSheet.setFontLine('none');
  teacherSheet.setBorder(false, false, false, false, false, false);
  teacherSheet.setHorizontalAlignment('center');
  teacherSheet.setVerticalAlignment('middle');
  teacherSheet.setWrap(true);

  // sets first row to bold
  sheet.getSheetByName('Teachers').getRange(1, 1, 1, teacherSheet.getLastColumn()).setFontWeight('bold');
}

// Jessica - gets all teacher info and sorts teachers alphabetically
function getTeachersArray() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let teacherSheet = sheet.getSheetByName("Teachers");
  let teacherValues = teacherSheet.getDataRange().getValues();

  // list of all teachers
  let teachersArray = [];

  // add every teacher's info to teachersArray
  for(let i = 1; i < teacherSheet.getLastRow(); i++) {
    let row = teacherValues[i]

    // removes timestamp & email
    row.splice(0, 2);

    // adds email to end of row
    row.push(teacherValues[i][1]);

    teachersArray.push(row);
  }

  return teachersArray;
}

// Angela - sorts the teachers sheet by the teacher's name
function sortTeachersSheet() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let teacherSheet = sheet.getSheetByName("Teachers");
  let teachersArray = teacherSheet.getDataRange().getValues();

  // removes first row
  teachersArray.splice(0, 1);

  // sorts the array
  teachersArray.sort(compareTeacherNames);

  // puts values in spreadsheet
  teacherSheet.getRange('A2:I' + (teachersArray.length + 1)).setValues(teachersArray);
}

// compares the values in the third column of the teacher sheet for the sortTeachersSheet method
function compareTeacherNames(a, b) {
  if (a[2] === b[2]) {
    return 0;
  }
  else {
    return (a[2] < b[2]) ? -1 : 1;
  }
}

// Jessica - organize teachers by prep periods
function organizedTeacherPeriods(teachersArray) {
  let organizedTeachers = [[], [], [], []];
  let period;

  //get list of teachers
  for(let i = 0; i < teachersArray.length; i++) {
    period = teachersArray[i][1];
    organizedTeachers[period - 1].push(teachersArray[i][0]);
  }

  return organizedTeachers;
}

// Angela - returns index of teacher being searched for
function binarySearchTeacher(teachersArray, teacher) {
  let first = 0;
  let last = teachersArray.length - 1;
  let midpoint;
  let index = -1;
  let found = false;

  while (!found && first <= last) {
    midpoint = Math.round((first + last) / 2);

    if (teacher == teachersArray[midpoint][0]) {
      found = true;
      index = midpoint;
    } else if (teacher > teachersArray[midpoint][0]) {
      first = midpoint + 1;
    } else {
      last = midpoint - 1;
    }
  }

  return index;
}

// Jessica + Angela - returns the teacher's prep period
function prepPeriod(teachersArray, teacher) {
  let index = binarySearchTeacher(teachersArray, teacher);

  if (index == -1) {
    return null;
  } else {
    return teachersArray[index][1];
  }
}

// Angela - returns the teacher's preferred duty
function preferredDuty(teachersArray, teacher) {
  let index = binarySearchTeacher(teachersArray, teacher);

  if (index == -1) {
    return null;
  } else {
    return teachersArray[index][2];
  }
}

// Angela - returns the teacher's unpreferred duty
function unpreferredDuty(teachersArray, teacher) {
  let index = binarySearchTeacher(teachersArray, teacher);

  if (index == -1) {
    return null;
  } else {
    return teachersArray[index][3];
  }
}

// Kiki + Angela - returns the teacher's contractual status
function getContractualStatus(teachersArray, teacher) {
  let index = binarySearchTeacher(teachersArray, teacher);

  if (index == -1) {
    return null;
  } else {
    return teachersArray[index][4];
  }
}

// Kiki + Angela - returns the teacher's number of shifts completed
function getActualShiftsCompleted(teachersArray, teacher) {
  let index = binarySearchTeacher(teachersArray, teacher);

  if (index == -1) {
    return null;
  } else {
    return teachersArray[index][6];
  }
}

// Kiki + Angela - returns the teacher's ratio
function getRatio(teachersArray, teacher) {
  let index = binarySearchTeacher(teachersArray, teacher);

  if (index == -1) {
    return null;
  } else {
    return teachersArray[index][6] / teachersArray[index][4];
  }
}

// Angela - returns the teacher's email
function getEmail(teachersArray, teacher) {
  let index = binarySearchTeacher(teachersArray, teacher);

  if (index == -1) {
    return null;
  } else {
    return teachersArray[index][8];
  }
}

// Angela - returns true if teacher wants emails, returns false otherwise
function wantsEmails(teachersArray, teacher) {
  let index = binarySearchTeacher(teachersArray, teacher);

  if (index == -1) {
    return null;
  } else if (teachersArray[index][5] == 'Yes') {
    return true;
  } else {
    return false;
  }
}

// Jessica + Angela - check if prep period matches with actual period
function isPrep(actualPrepPeriod, wantedPrep, day) {
  // actualPrepPeriod: block A, B, C, D - value is unchanging based on day 1 or 2
  // wantedPrep: value is changing based on day 1 or 2

  // day 1
  if (day == 1) {
    if (actualPrepPeriod == wantedPrep) {
      return true;
    }
    return false;

  // day 2
  } else {
    
    // if the blocks are the same on a day 2, it's not actually the same time period
    if (actualPrepPeriod == wantedPrep) {
      return false;
    }

    // if both are in the morning, then they are during the same time period
    if (actualPrepPeriod < 3 && wantedPrep < 3) {
      return true;
    }

    // if both are in the afternoon, then they are during the same time period
    if (actualPrepPeriod > 2 && wantedPrep > 2) {
      return true;
    }

    return false;
  }
}

// Jessica - check if available for lunch duty
function isAvailableForLunchDuty(period, day) {
  //check if days align
  if(day == 1) {
    if(period == 2 || period == 3) {
      return true;
    }
  }
  else if(day == 2){
    if(period == 1 || period == 4) {
      return true;
    }
  }
  return false;
}

// Angela - returns true if the shift parameter is the teacher's preferred shift; otherwise, returns false
function isPreferredDuty(teachersArray, teacher, shift) {
  let preferred = preferredDuty(teachersArray, teacher, shift);

  if (preferred == null) {
    return null;
  } else if (preferred == shift) {
    return true;
  } else {
    return false;
  }
}

// Angela - returns true if the shift parameter is the teacher's unpreferred shift; otherwise, returns false
function isUnpreferredDuty(teachersArray, teacher, shift) {
  let unpreferred = unpreferredDuty(teachersArray, teacher, shift);

  if (unpreferred == null) {
    return null;
  } else if (unpreferred == shift) {
    return true;
  } else {
    return false;
  }
}

// Kiki - returns true if the teacher needs to do more shifts
function needsMoreShifts(teachersArray, teacher, averageRatio) {
  let ratio = getRatio(teachersArray, teacher);

  if (ratio < averageRatio) {
    return true;
  } else {
    return false;
  }
}

// Kiki + Angela - returns the average ratio
function getAverageRatio(teachersArray) {
  let avg = 0;

  for (let i = 0; i < teachersArray.length; i++) {
    avg += teachersArray[i][6] / teachersArray[i][4];
  }

  avg /= teachersArray.length;
  return avg;
}

// Angela - returns the periods whose teachers with preps are available
function getAvailablePeriods(currentPeriod, currentDay) {
  let availablePeriods = [];

  // if it's at lunch
  if (currentPeriod == 0) {

    // cycles through all 4 periods
    for (let i = 1; i <= 4; i++) {
      if (isAvailableForLunchDuty(i, currentDay)) {
        availablePeriods.push(i);
      }
    }

  // otherwise, it's during a class
  } else {
    for (let i = 1; i <= 4; i++) {
      if (isPrep(i, currentPeriod, currentDay)) {
        availablePeriods.push(i);
        break;
      }
    }
  }

  return availablePeriods;
}

// Angela - filters the organizedTeachers array by only keeping the rows whose teachers are available
function filterbyAvailability(organizedTeachers, availablePeriods) {
  let availableTeachers = [];

  for (let i = 0; i < availablePeriods.length; i++) {
    availableTeachers.push([]);
    availableTeachers[i] = organizedTeachers[availablePeriods[i] - 1].slice();
  }

  return availableTeachers;
}

// Angela - filters the organizedTeachers array by removing any teacher who does not have the parameter duty as their preferred duty
function filterByPreferredDuty(organizedTeachers, teachersArray, duty) {
  let currentTeacher;
  let preferredTeachers = [];

  // adds the appropriate number of empty rows
  for (let i = 0; i < organizedTeachers.length; i++) {
    preferredTeachers.push([]);
  }

  // checks if each teacher has the preferred duty, adds them if so
  for (let i = 0; i < organizedTeachers.length; i++) {
    for (let j = 0; j < organizedTeachers[i].length; j++) {
      currentTeacher = organizedTeachers[i][j];

      if (isPreferredDuty(teachersArray, currentTeacher, duty)) {
        preferredTeachers[i].push(currentTeacher);
      }
    }
  }

  return preferredTeachers;
}

// Angela - filters the organizedTeachers array by removing any teacher who has the parameter duty as their unpreferred duty
function filterByUnpreferredDuty(organizedTeachers, teachersArray, duty) {
  let currentTeacher;
  let unpreferredTeachers = [];

  // adds the appropriate number of empty rows
  for (let i = 0; i < organizedTeachers.length; i++) {
    unpreferredTeachers.push([]);
  }

  // checks if each teacher has the preferred duty, adds them if so
  for (let i = 0; i < organizedTeachers.length; i++) {
    for (let j = 0; j < organizedTeachers[i].length; j++) {
      currentTeacher = organizedTeachers[i][j];

      if (!isUnpreferredDuty(teachersArray, currentTeacher, duty)) {
        unpreferredTeachers[i].push(currentTeacher);
      }
    }
  }

  return unpreferredTeachers;
}

// Kiki - gets a 2D array of teachers needing more shifts
function getTeachersNeedingShifts(organizedTeachers, teachersArray, averageRatio) {
  let teachersNeedingShifts = [];

  for (let i = 0; i < organizedTeachers.length; i++) {
    let periodTeachers = organizedTeachers[i];
    let needingShiftsInPeriod = [];

    
    for (let j = 0; j < periodTeachers.length; j++) {
      let teacher = periodTeachers[j];
      if (needsMoreShifts(teachersArray, teacher, averageRatio)) {
        needingShiftsInPeriod.push(teacher);
      }
    }
    
    teachersNeedingShifts.push(needingShiftsInPeriod);
  }

  return teachersNeedingShifts;
}

// Angela - counts total number of elements in a 2D array
function countTotalElements(array) {
  let count = 0;

  for (let i = 0; i < array.length; i++) {
    if (array[i].length != undefined) {
      count += array[i].length;
    }
  }

  return count;
}

// Angela
// parameter: 2D array (with two rows)
// returns: an array containing the index of a random element in the 2D array
function randomElement(array) {
  let totalElements = countTotalElements(array);
  let randomNum = Math.floor(Math.random() * totalElements);
  let coordinates = [];

  // if randomNum corresponds to an element in the first row
  if (randomNum < array[0].length) {
    coordinates.push(0);
    coordinates.push(randomNum);

  // if randomNum corresponds to an element in the second row
  } else {
    coordinates.push(1);
    coordinates.push(randomNum - array[0].length);
  }

  return coordinates;
}

// Angela - removes specified element from the 2D array
// "row" is the row in which the element is found
function removeElement(array, row, element) {
  for (let j = 0; j < array[row].length; j++) {
    if (element == array[row][j]) {
      array[row].splice(j, 1);
    }
  }

  return array;
}

// Angela - searches for the teacher's shifts in the specified month
// returns array:
// [[month, year, duty, day of week, day 1 or 2, time], [dates]]
function searchForTeacher(teacher, month, year) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let schedule = sheet.getSheetByName(month + " " + year).getRange('A4:I23').getValues();
  let values = [[month, year, '', '', '', ''], []];
  let teacherFound = false;

  // searches for teacher
  for (let currentRow = 0; currentRow < schedule.length; currentRow++) {
    for (let dayOfWeek = 0; dayOfWeek < 5; dayOfWeek++) {

      // looks for teacher's name in schedule
      if (teacher.toLowerCase() == schedule[currentRow][dayOfWeek + 4].toLowerCase()) {
        values[0][2] = schedule[Math.floor(currentRow / 2) * 2][0];
        values[0][3] = dayOfWeek;
        values[0][4] = schedule[currentRow][2];
        values[0][5] = schedule[Math.floor(currentRow / 2) * 2][1];

        teacherFound = true;
        break;
      }
    }
  }

  // converts month from string to number
  let monthNum;
  let months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  for (let i = 1; i <= months.length; i++) {
    if (month == months[i - 1]) {
      monthNum = i;
      break;
    }
  }

  let weekday = values[0][3];
  let day = values[0][4];

  // adds dates if teacher was found in schedule
  if (teacherFound) {
    let dates = getWeekdayDates(year, monthNum, weekday + 1, day);

    for (let j = 0; j < dates.length; j++) {
      values[1].push(dates[j]);
    }
  }

  // converts weekday from number to string
  switch (weekday) {
    case 0: weekday = 'Monday';    break;
    case 1: weekday = 'Tuesday';   break;
    case 2: weekday = 'Wednesday'; break;
    case 3: weekday = 'Thursday';  break;
    case 4: weekday = 'Friday';    break;
  }
  values[0][3] = weekday;

  return values;
}

// Angela -  on new searching sheet, a teacher can enter their name and their shift for the month pops up
function searchForShifts() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let search = sheet.getSheetByName('Search').getDataRange().getValues();
  let origSearchLen = search.length;
  let teacher = capitalizeFirstLetter(search[0][1]);
  let month = capitalizeFirstLetter(search[1][1]);
  let year = search[2][1];

  // searches for scheduled shifts
  let searchValues = searchForTeacher(teacher, month, year);

  // gets dates of duty
  let dates = searchValues[1];

  // empties all of the old dates
  search.splice(10, search.length - 10);

  // adds new dates
  for (let i = 0; i < dates.length; i++) {
    search.push([dates[i], '']);
  }

  // if there are less dates than the original search sheet, add new rows with empty strings
  if (search.length < origSearchLen) {
    for (let i = 0; i < (origSearchLen - search.length + 1); i++) {
      search.push(['', '']);
    }
  }

  // sets searched values into array
  search[4][1] = searchValues[0][2];
  search[5][1] = searchValues[0][3];
  search[6][1] = searchValues[0][4];
  search[7][1] = searchValues[0][5];

  console.log(search);

  // sets values in sheet
  sheet.getSheetByName('Search').getRange('A1:B' + search.length).setValues(search);

  // adds cell borders to dates

  // first, gets rid of borders for cells below dates
  sheet.getSheetByName('Search').getRange('A10:A' + search.length).setBorder(false, false, false, false, false, false);

  // adds borders for cells with dates
  sheet.getSheetByName('Search').getRange('A10:A' + (dates.length + 10)).setBorder(true, true, true, true, true, true);
}

// Angela - capitalizes the first letter of the parameter str
function capitalizeFirstLetter(str) {
  str.toLowerCase();
  str = str[0].toUpperCase() + str.substring(1);

  return str;
}

// Kiki - days of the week calculator 
function getDayOfWeekOccurrences(year, month, dayOfWeek) {
    const firstDayOfMonth = new Date(year, month - 1, 1);
    const lastDayOfMonth = new Date(year, month, 0);
    const daysInMonth = lastDayOfMonth.getDate();
    const firstDayOfWeek = firstDayOfMonth.getDay();
    const targetDayOfWeek = (dayOfWeek - firstDayOfWeek + 7) % 7;

    const occurrences = [];
    for (let day = targetDayOfWeek + 1; day <= daysInMonth; day += 7) {
        occurrences.push(day);
    }

    return occurrences;
}

// Jessica - get array of holidays
function getHolidays() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input");
  let holidayInputs = sheet.getDataRange().getValues();
  // 2d array of all holidays
  let holidays = [];

  for(let i = 5; i < sheet.getLastRow(); i++) {
    let year = holidayInputs[i][0];
    let month = holidayInputs[i][1];
    let day = holidayInputs[i][2];
    let type = holidayInputs[i][3];
    let row = [year, month, day, type];

    holidays.push(row);
  }

  return holidays;
}

// Jessica - get dates of a given weekday without holidays
function getNoHolidayOccurrences(year, month, dow) {
  let holidays = getHolidays();
  let occurrences = getDayOfWeekOccurrences(year, month, dow);
  let originalOccurrences = [];

  for(let i = 0; i < occurrences.length; i++) {
    originalOccurrences.push(occurrences[i]);
  }

  // for each date...
  for(let i = 0; i < originalOccurrences.length; i++) {
    let day = originalOccurrences[i];
    for(let j = 0; j < holidays.length; j++) {
      // get year, month, day of current holiday
      let hYear = holidays[j][0];
      let hMonth = holidays[j][1];
      let hDay = holidays[j][2];
      if(year == hYear && month == hMonth && day == hDay) {
        occurrences.splice(occurrences.indexOf(day), 1);
        break;
      }
    }
  }
  
  return occurrences;

}

// Jessica - get day1s or 2s of specific weekday without holidays
function organizedNoHolidayOccurrences(year, month, dow, day){
  let day1s = [];
  let day2s = [];

  let noHolidayOccurrences = getNoHolidayOccurrences(year, month, dow);
  for(let i = 0; i < noHolidayOccurrences.length; i++) {
    if(noHolidayOccurrences[i] %2 == 0) {
      day2s.push(noHolidayOccurrences[i]);
    } else {
      day1s.push(noHolidayOccurrences[i]);
    }
  }

  if(day == 1) {
    return day1s;
  } else {
    return day2s;
  }

}

// Jessica - get start and end dates of a school year
function getStartEndDates() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input");
  let dates = sheet.getDataRange().getValues();
  let startEndDates = [];
  // year, month, day format

  for(let i = 1; i <= 2; i++) {
    let year = dates[i][1];
    let month = dates[i][2];
    let day = dates[i][3];
    let row = [year, month, day];

    startEndDates.push(row);    
  }

  return startEndDates;
}

// Jessica - get day 1s or 2s of a month in a year
function getMonthDays(year, month, day, holiday) {
  // all day 1 dates by day of the week
  let day1s = [[], [], [], [], []];
  // all day 2 dates by day of the week
  let day2s = [[], [], [], [], []];

  let holidays = getHolidays();
  for (let dow = 1; dow <= 5; dow++) {
      let occurrences = getDayOfWeekOccurrences(year, month, dow);
      // keep dates without holidays in array
      let orignalOccurrences = [];

      for(let i = 0; i < occurrences.length; i++) {
        orignalOccurrences.push(occurrences[i]);
      }
      // if want dates with holiday types...
      if(holiday == true) {
        for(let i = 0; i < holidays.length; i++) {
          // check if year and month match the holiday
          if(holidays[i][0] == year && holidays[i][1] == month) {
            // go through all dates of a weekday in a month
            for(let j = 0; j < occurrences.length; j++) {
              // if holiday date and date match, add holiday type 
              if(holidays[i][2] == occurrences[j]) {
                occurrences[j] = " " + holidays[i][3] + " " + occurrences[j];
              }
            }
          }
        }
      }
      for (let i = 0; i < occurrences.length; i++) {
        // organize dates by day 1 and 2
        if(orignalOccurrences[i] %2 == 0) {
          day2s[dow-1].push(occurrences[i]);
        }
        else {
          day1s[dow-1].push(occurrences[i]);
        }
      }
  }

  if(day == 1) {
    return day1s;
  }
  else {
    return day2s;
  }
}

// Jessica - get dates for specific weekday for day 1 or 2
function getWeekdayDates(year, month, weekday, day) {
  // array of all the days of the month for day 1 or 2
  let allWeekdays = getMonthDays(year, month, day, true);
  // array of all days of month for specific weekday 
  let weekdayDays = allWeekdays[weekday - 1];
  //returns the numbers of the dates in an array
  return weekdayDays;
}

//Jessica - generate wanted month's schedule
function generateNewSchedule() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let generator = sheet.getSheetByName("Generate New Schedule").getDataRange().getValues();
  let month = generator[0][1];
  let year = generator[1][1];

  //capitalize first letter of month
  month = month.toLowerCase();
  month = month[0].toUpperCase() + month.substring(1);

  //name of new sheet
  let name = month + " " + year;

  let template = sheet.getSheetByName("Schedule Template");
  template.copyTo(sheet).setName(name);

  let currentSheet = sheet.getSheetByName(name).getDataRange().getValues();
  currentSheet[0][0] = month.toUpperCase();
  currentSheet[0][1] = year;


  // month as a num
  let monthNum = 0;
  // convert month string to num
  if(month == "January") {
    monthNum = 1;
  }
  else if(month == "February") {
    monthNum = 2;
  }
  else if(month == "March") {
    monthNum = 3;
  }
  else if(month == "April") {
    monthNum = 4;
  }
  else if(month == "May") {
    monthNum = 5;
  }
  else if(month == "June") {
    monthNum = 6;
  }
  else if(month == "July") {
    monthNum = 7;
  }
  else if(month == "August") {
    monthNum = 8;
  }
  else if(month == "September") {
    monthNum = 9;
  }
  else if(month == "October") {
    monthNum = 10;
  }
  else if(month == "November") {
    monthNum = 11;
  }
  else {
    monthNum = 12;
  }

  startEndDates = getStartEndDates();

  console.log(startEndDates);
  let startYear = startEndDates[0][0];
  let startMonth = startEndDates[0][1];
  let startDay = startEndDates[0][2];
  let endYear = startEndDates[1][0];
  let endMonth = startEndDates[1][1];
  let endDay = startEndDates[1][2];

  let start = false;
  let end = false;

  if(year ==  startYear && monthNum == startMonth) {
    start = true;
  } else if (year == endYear && monthNum == endMonth) {
    end = true;
  }

  // text for day 1 or day 2
  let day = "";
  // for both day 1 and day 2s
  for (let d = 1; d <= 2; d++) {
    // go through the weekday columns
    for(let c = 4; c <= 8; c++) {
      // get all day 1s or 2s
      let days = getMonthDays(year, monthNum, d, true);
      // get current weekday's days
      let currentDays = days[c-4];
      // day text
      if(d == 1) {
        day = "DAY 1: ";
      }
      else {
        day = "DAY 2: ";
      }
      // dates without holiday type text
      let noHolidayDays = getMonthDays(year, monthNum, d, false);
      let currentNoHDays = noHolidayDays[c-4];

      // if month wanted is a start month of school year...
      if(start == true) {
        // compare each day to the start date
        for(let i = 0; i < currentNoHDays.length; i++) {
          // cutoff the days before as soon as it hits the start day
          if(currentNoHDays[i] < startDay) {
            currentDays = currentDays.slice(i+1);
          }
        }
      }

      // if month wanted is last month of school year...
      if(end == true) {
        for(let i = 0; i < currentNoHDays.length; i++) {
          // cutoff the days after as soon as it passes the end day
          if(currentNoHDays[i] > endDay) {
            currentDays = currentDays.slice(0, i);
            break;
          }
        }
      }

      currentSheet[d-1][c] = day + currentDays.toString();
    }
  }

  sheet.getSheetByName(name).getRange("E1:I2").setWrap(true).setVerticalAlignment("middle");

  sheet.getSheetByName(name).getDataRange().setValues(currentSheet);

  generateSchedule(name);

  updateNumShiftsCompleted(name, year, monthNum);
}

// Jessica - add to teacher num shifts completed
function updateNumShiftsCompleted(sheetName, year, monthNum) {
  // let sheetName = "March 2024";
  // let year = 2024
  // let monthNum = 3
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  // current month schedule
  let schedule = sheet.getSheetByName(sheetName);
  // month schedule values
  let scheduleValues = schedule.getRange('E4:I23').getValues();
  // sheet with teacher info
  let teacherSheet = sheet.getSheetByName("Teachers");
  // get teacher sheet's values
  let teacherValues = teacherSheet.getDataRange().getValues();

  teachersArray = getTeachersArray();

  // dates of weekdays organized by day1 and 2s excluding holidays
  let occurrences;
  // for each column of teachers in schedule values
  for(let i = 0; i < scheduleValues[0].length; i++) {
    // for each row of teachers in schedule values
    for(let j = 0; j < scheduleValues.length; j++) {
      if((j+1) %2 != 0) {
        occurrences = organizedNoHolidayOccurrences(year, monthNum, i+1, 1);
      } else {
        occurrences = organizedNoHolidayOccurrences(year, monthNum, i+1, 2);
      }
      let numShifts = occurrences.length;
      // get current teacher in schedule
      let teacher = scheduleValues[j][i];
      console.log(teacher);
      // get teacher's position in teachersArray
      let teacherPosition = binarySearchTeacher(teachersArray, teacher); 
      //console.log(teacherPosition);
      // +1 to skip after header
      // update num shifts on teacher sheet
      teacherValues[teacherPosition + 1][8] += numShifts;
    }
  }

  //console.log(teacherValues);
  teacherSheet.getDataRange().setValues(teacherValues);

  // console.log(occurences);
  // //return teacherList;
  // console.log(scheduleValues);
}

// kiki sends automated schedule to email
function sendScheduleByEmail() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  let dataRange = sheet.getDataRange();
  let data = dataRange.getValues();

  // Iterate over each row in the spreadsheet
  for (let i = 1; i < data.length; i++) {
    let teacher = data[i][2];
    let sendSchedule = data[i][7];
    if (sendSchedule.toUpperCase() === 'Yes') {
        let month = "May";
        let year = "2024";
        console.log('Teacher:', teacher);
        let teacherSchedule = searchForTeacher(teacher, month, year)
      if (teacherSchedule && teacherSchedule.length > 0) {
        let email = data[i][1];
        let subject = 'Your Schedule for ' + month + ' ' + year;
        let message = 'Here is your schedule for ' + month + ' ' + year + ':\n';
        message += 'Duty: ' + teacherSchedule[2] + '\n';
        message += 'Weekday: ' + teacherSchedule[3] + '\n';
        message += 'Day: ' + teacherSchedule[0] + '\n';
        message += 'Time: ' + teacherSchedule[1] + '\n';
        GmailApp.sendEmail(email, subject, message);
      } 
      else {
        console.error('Teacher schedule not found for ' + teacher);
      }
    }
  }
}
