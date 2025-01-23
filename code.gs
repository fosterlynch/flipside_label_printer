// onOpen is a google sheets specific term
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    { name: 'print selcted rows to labels', functionName: 'printSelectedLabel' },
    { name: 'print all rows to labels', functionName: 'printAllLabels' }
  ];
  spreadsheet.addMenu('Print Labels', menuItems);
};

function printSelectedLabel() {
  console.log("Starting print selected labels");
  const datetime = Utilities.formatDate(new Date(), "GMT-7", 'EEE, MMM d yyyy h:mm:ss a');


  // Get the Google Sheet with the data
  var ss = SpreadsheetApp.getActive().getSheetByName("sample sheet"); // sample sheet comes from what the tab is named

  // Get the data from the sheet
  var selection = ss.getSelection();

  if (selection.getActiveRange().getA1Notation() == null) {
    ui.alert('No range selected');
    var list1data = ss.getDataRange().getValues();
    console.log("nothing selected");
  }

  else {
    console.log("active selection");
    var list1data = ss.getRange(selection.getActiveRange().getValues());
    console.log(list1data);
  }


  var originallist = ss.getDataRange().getValues();
  console.log(originallist)
}

function printAllLabels() {
  // const datetime = Utilities.formatDate(new Date(), "GMT-7", 'MM-dd-yyyy\'T\'aHH:mm:ss\'Z\'');
  const datetime = Utilities.formatDate(new Date(), "GMT-7", 'EEE, MMM d yyyy h:mm:ss a');

  console.log("Starting print all labels");
  // Get the Google Sheet with the data
  var ss = SpreadsheetApp.getActive().getSheetByName("sample sheet"); // sample sheet comes from what the tab is named
  // Get the data from the sheet
  var selection = ss.getSelection();
  // ui = SpreadsheetApp.getUi();

  // if a user has a cell clicked, i think that changes the active range
  // going to bypass control flow for now
  if (selection.getActiveRange().getA1Notation() == null) {
    ui.alert('No range selected');
    var data = ss.getDataRange().getValues();
    // console.log("here");
  }
  else {
    var range = ss.getRange(selection.getActiveRange().getA1Notation());
    // console.log("no, here!");
  }
  var list1data = ss.getDataRange().getValues();
  const uiNames = list1data[0];
  const columnNames = list1data[1];
  // console.log(columnNames);
  for (i = 1; i < list1data.length; i++) {

    // Get the product label template
    // var folderID = 'flipside/templates'; //todo, change to nested structure eventually
    // var doc = DocumentApp.openById(template_id);
    // var template = doc.getBody();

    // // Replace the placeholders in the template with the data from the sheet
    // var dir = DriveApp.getFolderById(DriveApp.getFoldersByName(folderID).next().getId());

    // grabbing global values
    var item_name = getValueByName(ss, "Item", i);
    var make = getValueByName(ss, "Make", i);
    var model = getValueByName(ss, "Model", i);
    var price = getValueByName(ss, "Sell Price", i);
    var condition = getValueByName(ss, "Condition", i);
    var item_type = getValueByName(ss, "Item Type", i);

    if (make != "" && price == "") {
      console.warn(`entry ${make} ${model} has no price point, skipping item`)
      continue; // if price is blank, we are skipping
    }

    if ((price == "") || (price == null) || (price == "Sell Price")) { // these checks are here to account for spreadsheet formatting for humans,
      continue; // if price is blank, we are skipping
    }

    console.log(item_name)
    switch (item_type) {
      case "M":
        continue;
      // var template_id = DriveApp.getFilesByName("miscellaneous").next().getId();
      // console.log("generating label", item_name, price);
      // var copyId = DriveApp.getFileById(template_id).makeCopy(item_name).getId();
      // console.log("log: copyID", copyId);

      // // Open the temporary document
      // var copyDoc = DocumentApp.openById(copyId);
      // // Get the document’s body section
      // var label = copyDoc.getBody();
      // // Replace the placeholders in the template

      // break;
      case "G":
        console.log(`starting spreadsheet extraction for ${item_name}`);
        var template_id = DriveApp.getFilesByName("guitar_template").next().getId();
        console.log("generating label", item_name, price);
        var copyId = DriveApp.getFileById(template_id).makeCopy(item_name).getId();
        console.log("log: copyID", copyId);

        // Open the temporary document
        var copyDoc = DocumentApp.openById(copyId);
        // Get the document’s body section
        var label = copyDoc.getBody();
        // Replace the placeholders in the template

        var origin = getValueByName(ss, "Origin", i);
        var serial = getValueByName(ss, "Serial #", i);
        var scale = getValueByName(ss, "Item Type", i);
        var body_wood = getValueByName(ss, "Body Wood", i);
        var color = getValueByName(ss, "Color", i);
        var fretboard = getValueByName(ss, "Fretboard", i);
        var pickups = getValueByName(ss, "Pick ups", i);
        var includes_case = getValueByName(ss, "Case", i);
        var moded = getValueByName(ss, "Modded", i);

        label.replaceText("{ORIGIN}", origin);
        label.replaceText("{SERIAL}", serial);
        label.replaceText("{SCALE}", scale);
        label.replaceText("{BODY_WOOD}", body_wood);
        label.replaceText("{COLOR}", color);
        label.replaceText("{FRETBOARD}", fretboard);
        label.replaceText("{PICKUPS}", pickups);
        label.replaceText("{INCLUDES_CASE}", includes_case);
        label.replaceText("{MODDED}", moded);
        break;
      case "A":
        continue;
      // var template_id = DriveApp.getFilesByName("amp").next().getId();
      // console.log("generating label", item_name, price);
      // var copyId = DriveApp.getFileById(template_id).makeCopy(item_name).getId();
      // console.log("log: copyID", copyId);

      // // Open the temporary document
      // var copyDoc = DocumentApp.openById(copyId);
      // // Get the document’s body section
      // var label = copyDoc.getBody();

      // // Replace the placeholders in the template
      // var channel = getValueByName(ss, "Channel", i);
      // var power_tubes = getValueByName(ss, "Origin", i);
      // var speaker = getValueByName(ss, "Serial #", i);
      // var fs = getValueByName(ss,"Item Type",i);
      // var effects = getValueByName(ss,"Body Wood",i);
      // var cover = getValueByName(ss,"Fretboard",i);

      // label.replaceText("{CHANNEL}", channel);
      // label.replaceText("{POWER_TUBES}", power_tubes);
      // label.replaceText("{SPEAKER}", speaker);
      // label.replaceText("{FS}", fs);
      // label.replaceText("{EFFECTS}", effects);
      // label.replaceText("{COVER}", cover);

      // break;

      case "P":
        continue;
      // var template_id = DriveApp.getFilesByName("pedal").next().getId();
      // console.log("generating label", item_name, price);
      // var copyId = DriveApp.getFileById(template_id).makeCopy(item_name).getId();
      // console.log("log: copyID", copyId);

      // // Open the temporary document
      // var copyDoc = DocumentApp.openById(copyId);
      // // Get the document’s body section
      // var label = copyDoc.getBody();

      // // Replace the placeholders in the template
      // var box = getValueByName(ss, "Box", i);
      // var power = getValueByName(ss, "Power", i);
      // label.replaceText("{BOX}", box);
      // label.replaceText("{POWER}", power);
      // break;

    }

    label.replaceText("{MAKE}", make);
    label.replaceText("{MODEL}", model);
    label.replaceText("{PRICE}", price);
    label.replaceText("{CONDITION}", condition);
    console.log(":))))");
    // copyDoc.getBody().setAttributes({"PAGE_WIDTH":52,"PAGE_HEIGHT":152}),
    // ui.alert("saved PDF for item", make, model)
    copyDoc.saveAndClose();
    // console.log("copy doc", copyDoc);
    docToPDF(copyDoc, datetime)
    deleteFileByID(copyId)
  }
}

function docToPDF(docfile, datetime) {

  // get Google Drive folder
  var folder_ID = DriveApp.getFoldersByName('COGS label printer').next().getId();
  var parentFolder = DriveApp.getFolderById(folder_ID); //add this line...
  console.log(`datetime is ${datetime}`);
  var folder, folders = DriveApp.getFoldersByName("label_exports " + datetime);

  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = parentFolder.createFolder("label_exports " + datetime); //edit this line
  }



  // get file content as PDF blob
  var pdfBlob = docfile.getAs('application/pdf');
  pdfBlob.setName(docfile.getName() + "_label_.pdf")
  // create new PDF file in Google Drive folder
  folder.createFile(pdfBlob);
  // console.log(folder)

  return "Thank you, your file was uploaded successfully!";
}

function deleteFileByID(fileId) {
  var file = Drive.Files.get(fileId);
  if (file.mimeType === MimeType.FOLDER) {
    // possibly ask for confirmation before deleting this folder
  }
  Drive.Files.remove(file.id); // "remove" in Apps Script client library, "delete" elsewhere
}

function getValueByName(sheet, colName, row) {
  var data = sheet.getDataRange().getValues();
  var col = data[1].indexOf(colName);
  if (col != -1) {
    return data[row - 1][col];
  }
}


