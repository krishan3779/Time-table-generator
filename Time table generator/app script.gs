var progress = 0; //record procceed line
var course = {}; //record generated courses
var classDetail = []; //record generated classes
var colourRegister = 0; //record used colour
var colour = ['#ffbf91', '#ebff91', '#aeff91', '#aeff91', '#8fd0ff', '#9d9cff', '#d99cff', '#fce7d7', '#e9fcd7', '#d7fce0', '#d7fbfc', '#d7e5fc', '#e3d7fc'];
var file = SpreadsheetApp.openById("1N0C3IK1Dx_xJNQpY-tVwkW2P2A8tU8uGeJ2lWbAsVNg");
var timetablePage = file.getSheetByName("timetable");
var dataPage = file.getSheetByName("data");

/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  var menu = [{ name: 'Create a new timetable', functionName: 'createTimeTable' },
  { name: 'Create a new timetable for a course and form', functionName: 'createTimeTableWithRestriction' },
  { name: 'Initialize timetable', functionName: 'initializeTimetable' }];
  SpreadsheetApp.getActive().addMenu('Timetable', menu);
}


function searchClashingClass(week, startTime) {
  startTime = time2Number(startTime);
  var i = -1;
  classDetail.forEach(function (e, index) {
    if (e.week == week && time2Number(e.endTime) > startTime) {
      i = index;
    }
  });

  return i;
}

function registerNewCourse(registerClassName) {
  if (!isCourseExist(registerClassName)) {
    course[registerClassName] = {
      colour: colour[colourRegister]
    };
    ++colourRegister;
  }
}

function isCourseExist(alias) {
  if (course[alias] == undefined)
    return false;
  return true;
}


function drawClassOnTimetable(className, classNameArray, week, startTime, endTime, form) {
  if (!isCourseExist(className))
    registerNewCourse(className);
  week = week2Number(week);
  var startCell = (startTime.hr - 9) * 4 + 2 + Math.floor(startTime.mins / 15);
  var cellLength = 0;
  if (startTime.hr == endTime.hr) { //class will end in same hour
    cellLength = Math.floor((endTime.mins - startTime.mins) / 15);
  } else {
    cellLength = (endTime.hr - startTime.hr) * 4 - Math.floor(startTime.mins / 15) + Math.floor(endTime.mins / 15);
  }

  var timeSlot = timetablePage.getRange(startCell, week + 1, cellLength, 1);
  timeSlot.clearFormat().clearContent().setBackground(course[className].colour).setBorder(true, true, true, true, false, false, "#000000", null).setHorizontalAlignment('center');
  var classNameUsedSpace = 1;
  if (classNameArray.length == 1) { //only one class is hold in this time slot
    timeSlot.getCell(1, 1).setValue(className);
    ++classNameUsedSpace;
  } else {
    //more than one class is hold in this time slot
    var numberOfClassName = Math.ceil(classNameArray.length / (cellLength - 2));  //calculate how many classes can be displayed in a cell
    for (var i = 0; i < classNameArray.length; i += numberOfClassName) {
      var text = classNameArray[i];
      for (var r = 1; r < numberOfClassName; r++) {
        if (i + r < classNameArray.length) {
          text += '/' + classNameArray[i + r];
        } else {
          break;
        }
      }
      timeSlot.getCell(classNameUsedSpace, 1).setValue(text);
      ++classNameUsedSpace;
    }
  }
  timeSlot.getCell(classNameUsedSpace, 1).setValue(form);
  timeSlot.getCell(cellLength, 1).setValue(startTime.hr + ':' + startTime.mins + ' - ' + endTime.hr + ':' + endTime.mins);

}

function week2Number(week) {
  if (week == "Mon" || week == "mon" || week == "一")
    return 1;
  if (week == "Tue" || week == "tue" || week == "二")
    return 2;
  if (week == "Wed" || week == "wed" || week == "三")
    return 3;
  if (week == "Thur" || week == "thur" || week == "四")
    return 4;
  if (week == "Fri" || week == "fri" || week == "五")
    return 5;
  if (week == "Sat" || week == "sat" || week == "六")
    return 6;
  if (week == "Sun" || week == "sun" || week == "日")
    return 7;
  throw new Error("行數：" + progress + " 輸入了錯誤的星期。");
}

function time2Number(time) {
  return parseInt(time.hr) * 60 + parseInt(time.mins);
}

function form2Level(form) {
  var firstChar = form.substr(0, 1);
  var level = 0;
  if (firstChar == "P")
    level += 3 + parseInt(form.substr(form.length - 1, form.length));
  else if (firstChar == "F")
    level += 9 + parseInt(form.substr(form.length - 1, form.length));
  else if (firstChar != "K")
    level = 100; //the form is not in KPF system, maybe is adult / private

  return level;
}

function getHour(hr) {
  hr = String(hr);

  var index = String(hr).indexOf(":") || String(hr).indexOf("：");
  if (index == -1)
    throw new Error("行數：" + progress + " 輸了錯誤的上課／下課時間。");
  return hr.substr(0, index);
}

function getMinutes(mins) {
  mins = String(mins);

  var index = String(mins).indexOf(":") || String(mins).indexOf("：");
  if (index == -1)
    throw new Error("行數：" + progress + " 輸了錯誤的上課／下課時間。");
  return mins.substr(index + 1, mins.length);
}

function createTimeTableWithRestriction() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var courseName = ui.prompt(
    '輸入課程名稱：',
    '',
    ui.ButtonSet.OK_CANCEL);

  var ok = courseName.getSelectedButton();

  if (ok == ui.Button.OK) {

    var formName = ui.prompt(
      '輸入年級限制：',
      '',
      ui.ButtonSet.OK_CANCEL);

    ok = formName.getSelectedButton();

    if (ok == ui.Button.OK) {
      courseName = courseName.getResponseText();
      formName = formName.getResponseText();
      createTimeTable(courseName, formName);
    }
  }
}

function initializeTimetable() {
  var content = timetablePage.getRange(2, 2, 44, 7);
  content.clearContent();
  content.setBorder(true, true, true, true, false, false, "#000000", null); //set border style of the entire timetable
  content.setBackground("#FFFFFF");

  for (var c = 0; c <= 6; c++) {
    for (var r = 0; r <= 10; r++) {
      var border = timetablePage.getRange(2 + r * 4, 2 + c, 4, 1);
      border.setBorder(true, true, true, true, false, false, "#000000", null); //set border style for each time slot
    }
  }
}

function createTimeTable(courseName, formName) {
  var clashingClassIndex;
  var content = timetablePage.getRange(2, 2, 44, 7);
  content.clearContent();
  content.setBorder(true, true, true, true, false, false, "#000000", null); //set border style of the entire timetable
  content.setBackground("#FFFFFF");

  for (var c = 0; c <= 6; c++) {
    for (var r = 0; r <= 10; r++) {
      var border = timetablePage.getRange(2 + r * 4, 2 + c, 4, 1);
      border.setBorder(true, true, true, true, false, false, "#000000", null); //set border style for each time slot
    }
  }

  while (true) {
    ++progress;
    var data = dataPage.getRange(progress, 1, 1, 15).getValues();
    var tutorialClassName = data[0][1];
    if (tutorialClassName == "") {
      break; //read a blank line, stop generating.
    } else {
      var week = data[0][8];

      var start = {
        hr: getHour(data[0][6]),
        mins: getMinutes(data[0][6])
      };

      var end = {
        hr: getHour(data[0][7]),
        mins: getMinutes(data[0][7])
      };

      if (courseName == undefined || (courseName == tutorialClassName && Math.abs(form2Level(formName) - form2Level(data[0][3])) <= 1)) {
        clashingClassIndex = searchClashingClass(week, start);
        if (clashingClassIndex == -1) { //first class should not be merged
          var tutorialClass = {
            className: tutorialClassName,
            classNameInArray: [tutorialClassName],
            week: week,
            startTime: start,
            endTime: end,
            form: data[0][3]
          };
          classDetail.push(tutorialClass);
        } else {
          if (classDetail[clashingClassIndex].className.match(tutorialClassName) == null) {
            classDetail[clashingClassIndex].className += '/' + tutorialClassName;
            classDetail[clashingClassIndex].classNameInArray.push(tutorialClassName);
          }
          if (time2Number(start) < time2Number(classDetail[clashingClassIndex].startTime))
            classDetail[clashingClassIndex].startTime = start;
          if (time2Number(end) > time2Number(classDetail[clashingClassIndex].endTime))
            classDetail[clashingClassIndex].endTime = end;
          if (classDetail[clashingClassIndex].form.match(data[0][3]) == null)
            classDetail[clashingClassIndex].form += '/' + data[0][3];
        }
      }
    }
  }

  classDetail.forEach(function (e) {
    drawClassOnTimetable(e.className, e.classNameInArray, e.week, e.startTime, e.endTime, e.form);
  });
}
