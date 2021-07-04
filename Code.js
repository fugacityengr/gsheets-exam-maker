function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu("Forms");
  menu.addItem("CREATE TEMPLATE", "createTemplate").addToUi();
  menu.addItem("CREATE FORM", "createForm").addToUi();
}

function createTemplate() {
  // Initialize spreadsheet
  var s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Reduce number of columns and rows displayed
  s.deleteColumns(17, 8);
  s.deleteRows(101, 900);
  s.setColumnWidth(1, 110);
  s.setColumnWidth(2, 400);
  s.setColumnWidths(3, 16, 110);

  // Freeze Question Column
  s.setFrozenColumns(2);
  s.setFrozenRows(3);

  // Format column A to be BOLD and CENTER
  s.getRange("A1:A101").setFontWeight("bold").setHorizontalAlignment("center");

  // Form Object Types
  s.getRange("A4").setValue("TYPE");
  s.getRange("A5:A101").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .setAllowInvalid(false)
      .requireValueInList(
        [
          "CHECKBOX",
          "CHOICE",
          "DATE",
          "LIST",
          "PAGE",
          "PARAGRAPH",
          "SECTION",
          "TEXT",
          "TIME",
        ],
        true
      )
      .build()
  );

  // Form Title
  s.getRange("A1").setValue("FORM TITLE");

  // Form Description
  s.getRange("A2").setValue("DESCRIPTION");

  // Folder ID
  s.getRange("A3").setValue("FOLDER ID:").setFontWeight("bold");

  // Generated Form Public URL
  s.getRange("C1")
    .setValue("FORM URL:")
    .setFontWeight("bold")
    .setHorizontalAlignment("right");

  // Question / Labels
  s.getRange("B4")
    .setValue("QUESTION")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Points
  s.getRange("C4")
    .setValue("POINTS")
    .setFontWeight("bold")
    .setBackground("#ffff66")
    .setHorizontalAlignment("center");

  // Options
  s.getRange("D4:H4")
    .setValue("OPTION")
    .setFontWeight("bold")
    .setBackground("#d4dee5")
    .setHorizontalAlignment("center");

  // Options Column Formatting
  s.getRange("D5:H101").setBackground("#d4dee5");

  // Points Column Formatting
  s.getRange("C5:C101").setBackground("#ffff66");
}

function createForm() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  body(s);
}

// Helper function to set question choices
function choiceMaker(sheet, rowNum, rowData, numColumns, question) {
  var arr = [];
  for (var ccc = 3; ccc < numColumns; ccc++) {
    var cu = 1 + ccc;
    var cellData = sheet.getRange(cr, cu, 1, 1).getValue();
    var cellColor = sheet.getRange(cr, cu, 1, 1).getBackground();
    if (cellData === "") continue;
    switch (cellColor) {
      case "#00ff00":
        arr.push(question.createChoice(rowData[rowNum][ccc], true));
        break;
      default:
        arr.push(question.createChoice(rowData[rowNum][ccc], false));
        break;
    }
  }
  question.setChoices(arr);
}

// Helper function to check if points were given and to set if available
function pointSetter(rowNum, rowData, question) {
  if (rowData[rowNum][2] !== "") {
    question.setPoints(rowData[rowNum][2]);
  }
}

// TODO: Break apart body function to separate functions handling each type of form object
// Main Function Call to Create Form
function body(s) {
  var r = s.getDataRange();
  var nr = r.getNumRows();
  var nc = r.getNumColumns();
  var d = r.getValues();

  // Get Drive Folder
  var fol = DriveApp.getFolderById(d[2][1]);

  // Create form with Form Title
  var fm = FormApp.create(d[0][1]);
  // Get the id of the created form object
  var id = fm.getId();
  // Open the form object
  var f = FormApp.openById(id);

  // Set Form Description
  f.setDescription(d[1][1]);
  // Set the Form as a Quiz Form
  f.setIsQuiz(true);

  // Get the Public URL of the Form and place on D1
  var ur = f.getPublishedUrl();
  s.getRange("D1").setValue(ur);

  // Get the id of the Google Form file in Google Drive
  var file = DriveApp.getFileById(id);
  // Add this file to the specified folder
  // By default, forms created are added to the root folder of Google Drive
  file.moveTo(fol);
  // !Deprecated method, replaced by moveTo
  // fol.addFile(file);
  // DriveApp.getRootFolder().removeFile(file);

  // Iterate over the rows
  for (var x = 4; x < nr; x++) {
    // Get form object type
    var i = d[x][0];

    switch (i) {
      case "":
        // Move on to the next cell
        continue;

      case "CHOICE":
        var q = f.addMultipleChoiceItem().setTitle(d[x][1]).setRequired(true);
        pointSetter(x, d, q);
        choiceMaker(s, x, d, nc, q);

        break;

      case "LIST":
        var q = f.addListItem().setTitle(d[x][1]).setRequired(true);
        pointSetter(x, d, q);
        choiceMaker(s, x, d, nc, q);

        break;

      case "CHECKBOX":
        var q = f.addCheckboxItem().setTitle(d[x][1]).setRequired(true);
        pointSetter(x, d, q);
        choiceMaker(s, x, d, nc, q);
        break;

      case "DATE":
        f.addDateItem().setTitle(d[x][1]).setRequired(true);

        break;

      case "PAGE":
        f.addPageBreakItem().setTitle(d[x][1]);

        break;

      case "PARAGRAPH":
        var q = f.addParagraphTextItem().setTitle(d[x][1]).setRequired(true);
        pointSetter(x, d, q);

        break;

      case "SECTION":
        f.addSectionHeaderItem().setTitle(d[x][1]);

        break;

      case "TEXT":
        var q = f.addTextItem().setTitle(d[x][1]).setRequired(true);
        pointSetter(x, d, q);

        break;

      case "TIME":
        var q = f.addTimeItem().setTitle(d[x][1]).setRequired(true);
        pointSetter(x, d, q);
        break;
    }
  }
}
