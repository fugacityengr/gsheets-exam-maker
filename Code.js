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

// TODO: Break apart body function to separate functions handling each type of form object
// Main Function Call to Create Form
function body(s) {
  var r = s.getDataRange();
  var nr = r.getNumRows();
  var nc = r.getNumColumns();
  var lr = s.getLastRow();
  var lc = s.getLastColumn();
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
    // Ranges used in getting values
    var cr = 1 + x;
    var ro = s.getRange(cr, 4, 1, 5);
    var op = ro.getValues();

    // TODO: Change if-else statements to switch
    // TODO: Continue rewriting logic from LIST

    switch (i) {
      case "":
        continue;

      case "CHOICE":
        var arr = [];

        var q = f.addMultipleChoiceItem();
        q.setTitle(d[x][1]).setRequired(true);

        if (d[x][2] !== "") {
          q.setPoints(d[x][2]);
        }

        for (var ccc = 3; ccc < nc; ccc++) {
          var cu = 1 + ccc;
          var cellData = s.getRange(cr, cu, 1, 1).getValue();
          var cellColor = s.getRange(cr, cu, 1, 1).getBackground();
          if (cellData === "") continue;
          switch (cellColor) {
            case "#00ff00":
              arr.push(q.createChoice(d[x][ccc], true));
              break;
            default:
              arr.push(q.createChoice(d[x][ccc], false));
              break;
          }
        }
        q.setChoices(arr);
        break;

      case "LIST":
        var arr = [];
        var q = f.addListItem();

        q.setTitle(d[x][1]).setRequired(true);

        if (d[x][2] !== "") {
          q.setPoints(d[x][2]);
        }

        for (var ccc = 3; ccc < nc; ccc++) {
          var cu = 1 + ccc;
          var cellData = s.getRange(cr, cu, 1, 1).getValue();
          var cellColor = s.getRange(cr, cu, 1, 1).getBackground();
          if (cellData === "") continue;
          switch (cellColor) {
            case "#00ff00":
              arr.push(q.createChoice(d[x][ccc], true));
              break;
            default:
              arr.push(q.createChoice(d[x][ccc], false));
              break;
          }
        }
        q.setChoices(arr);
        break;

      case "CHECKBOX":
        break;

      case "DATE":
        break;

      case "PAGE":
        break;

      case "PARAGRAPH":
        break;

      case "SECTION":
        break;

      case "TEXT":
        break;

      case "TIME":
        break;
    }

    if (i == "") {
      continue;
    } else if (i == "CHOICE") {
    } else if (i == "LIST") {
    } else if (i == "CHECKBOX") {
      var arr = [];

      if (d[0][11] == "YES") {
        var its = f.getItems();
        for (var w = 0; w < its.length; w += 1) {
          var ite = its[w];
          if (ite.getTitle() === "CHECKBOX") {
            var q = ite.asCheckboxItem().duplicate();
          }
        }
      } else {
        var q = f.addCheckboxItem();
      }

      q.setTitle(d[x][1]).setHelpText(d[x][2]).setRequired(true);

      if (d[x][3] !== "") {
        q.setPoints(d[x][3]);
      }

      for (var ccc = 8; ccc < nc; ccc++) {
        var cu = 1 + ccc;
        if (
          s.getRange(cr, cu, 1, 1).getValue() !== "" &&
          s.getRange(cr, cu, 1, 1).getBackground() === "#00ff00"
        ) {
          var q1 = q.createChoice(d[x][ccc], true);
          arr.push(q1);
        } else if (
          s.getRange(cr, cu, 1, 1).getValue() !== "" &&
          s.getRange(cr, cu, 1, 1).getBackground() !== "#00ff00"
        ) {
          var q1 = q.createChoice(d[x][ccc], false);
          arr.push(q1);
        }
      }

      q.setChoices(arr);

      if (d[x][4] !== "") {
        var correctFeedback = FormApp.createFeedback().setText(d[x][4]).build();
        q.setFeedbackForCorrect(correctFeedback);
      }
      if (d[x][5] !== "") {
        var incorrectFeedback = FormApp.createFeedback()
          .setText(d[x][5])
          .addLink(d[x][6], d[x][7])
          .build();
        q.setFeedbackForIncorrect(incorrectFeedback);
      }
    } else if (i == "GRID") {
      var arr1 = [];
      for (q = 0; q < op[0].length; q++) {
        if (op[0][q] !== "") {
          arr1.push(op[0][q]);
        }
      }
      var arr2 = [];
      for (q = 0; q < op[0].length; q++) {
        if (op[0][q] !== "") {
          arr2.push(op[0][q]);
        }
      }
      f.addGridItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setRequired(true)
        .setRows(arr1)
        .setColumns(arr2);
    } else if (i == "CHECKGRID") {
      var arr1 = [];
      for (q = 0; q < op[0].length; q++) {
        if (op[0][q] !== "") {
          arr1.push(op[0][q]);
        }
      }
      var arr2 = [];
      for (q = 0; q < op[0].length; q++) {
        if (op[0][q] !== "") {
          arr2.push(op[0][q]);
        }
      }
      f.addCheckboxGridItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setRequired(true)
        .setRows(arr1)
        .setColumns(arr2);
    } else if (i == "TEXT") {
      var q = f
        .addTextItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setRequired(true);
      if (d[x][3] !== "") {
        q.setPoints(d[x][3]);
      }
    } else if (i == "PARAGRAPH") {
      var q = f
        .addParagraphTextItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setRequired(true);
      if (d[x][3] !== "") {
        q.setPoints(d[x][3]);
      }
    } else if (i == "SECTION") {
      f.addSectionHeaderItem().setTitle(d[x][1]).setHelpText(d[x][2]);
    } else if (i == "PAGE") {
      f.addPageBreakItem().setTitle(d[x][1]).setHelpText(d[x][2]);
    } else if (i == "IMAGE1") {
      var img = UrlFetchApp.fetch(d[x][6]);
      f.addImageItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setImage(img)
        .setAlignment(FormApp.Alignment.CENTER)
        .setWidth(800);
    } else if (i == "IMAGE2") {
      var file = DriveApp.getFileById(d[x][6]);
      f.addImageItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setImage(file)
        .setAlignment(FormApp.Alignment.CENTER)
        .setWidth(800);
    } else if (i == "VIDEO") {
      f.addVideoItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setVideoUrl(d[x][6])
        .setAlignment(FormApp.Alignment.CENTER)
        .setWidth(800);
    } else if (i == "SCALE") {
      var q = f
        .addScaleItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setRequired(true)
        .setLabels(d[x][6], d[x][7])
        .setBounds(d[x][4], d[x][5]);
      if (d[x][3] !== "") {
        q.setPoints(d[x][3]);
      }
    } else if (i == "TIME") {
      var q = f
        .addTimeItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setRequired(true);
      if (d[x][3] !== "") {
        q.setPoints(d[x][3]);
      }
    } else if (i == "DATE") {
      var q = f
        .addDateItem()
        .setTitle(d[x][1])
        .setHelpText(d[x][2])
        .setRequired(true);
      if (d[x][3] !== "") {
        q.setPoints(d[x][3]);
      }
    } else if (i == "ACCEPTANCE") {
      var item = f.addMultipleChoiceItem();
      var goSubmit = item.createChoice(
        "YES",
        FormApp.PageNavigationType.SUBMIT
      );
      var goRestart = item.createChoice(
        "NO",
        FormApp.PageNavigationType.RESTART
      );
      item.setRequired(true);
      item.setTitle(d[x][1]);
      item.setHelpText(d[x][2]);
      item.setChoices([goSubmit, goRestart]);
    }
  } // End of principle for loop with x

  var iti = f.getItems();
  for (var y = 0; y < iti.length; y += 1) {
    var ito = iti[y];
    if (ito.getTitle() === "CHOICE") {
      f.deleteItem(ito);
    } else if (ito.getTitle() === "LIST") {
      f.deleteItem(ito);
    } else if (ito.getTitle() === "CHECKBOX") {
      f.deleteItem(ito);
    }
  }
} // End of entire scipt
