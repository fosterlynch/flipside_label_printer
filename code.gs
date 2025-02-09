// onOpen is a google sheets specific term
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    { name: 'BETA TESTING: print selcted rows to labels', functionName: 'printSelectedLabel' },
    { name: 'print all rows to labels', functionName: 'printAllLabels' }
  ];
  spreadsheet.addMenu('Print Labels', menuItems);
  const ui = SpreadsheetApp.getUi();
  //   const response = ui.prompt( // This could be helpful for letting people do a tutorial
  //     'Getting to know you',
  //     'May I know your name?',
  //     ui.ButtonSet.OK,
  // );
};

function status(update) {
  const htmlOutput = HtmlService
    .createHtmlOutput(
      `<p>${update}...</p>`,
    )
    .setWidth(250)
    .setHeight(300);
  return htmlOutput
}

function printSelectedLabel() {
  const ui = SpreadsheetApp.getUi();
  console.log("Starting print selected labels");
  ui.showModalDialog(status("Starting Print Selected Labels"), 'Printing Labels');


  const datetime = Utilities.formatDate(new Date(), "GMT-7", 'EEE, MMM d yyyy h:mm:ss a');


  // Get the Google Sheet with the data
  var ss = SpreadsheetApp.getActive().getSheetByName("COGS Test Sheet"); // sample sheet comes from what the tab is named
  var data = ss.getActiveRange().getValues();
  var columns = ss.getDataRange().getValues()[1];
  var selection = ss.getSelection();

  if (selection.getActiveRange().getA1Notation() == null) {
    ui.alert('No range selected');
    var list1data = ss.getDataRange().getValues();
    console.log("nothing selected");
  }

  else {
    console.log("active selection");
    var range = SpreadsheetApp.getActiveSpreadsheet().getRange(selection.getActiveRange().getA1Notation());
    var list1data = range.getValues();
  }

  const pdfids = [];
  for (i = 1; i <= list1data.length; i++) {
    console.log("starting iteration")
    console.log(i)

    var loading = HtmlService
      .createHtmlOutput(
        `<p>Processing ${i} of ${list1data.length}...</p>`,
      )
      .setWidth(250)
      .setHeight(300);
    ui.showModalDialog(loading, 'Printing Labels');

    // grabbing global values
    var item_name = getSelectedValueByName(data, columns, "Item", i);
    console.log(item_name);
    var make = getSelectedValueByName(data, columns, "Make", i);
    var model = getSelectedValueByName(data, columns, "Model", i);
    var price = getSelectedValueByName(data, columns, "Sell Price", i);
    var condition = getSelectedValueByName(data, columns, "Condition", i);
    var item_type = getSelectedValueByName(data, columns, "Item Type", i);
    var notes = getSelectedValueByName(data, columns, "Notes", i);

    console.log("getting values")
    console.log(item_name)
    if (make != "" && price == "") {
      console.warn(`entry ${make} ${model} has no price point, skipping item`)
      continue; // if price is blank, we are skipping
    }

    if ((price == "") || (price == null) || (price == "Sell Price")) { // these checks are here to account for spreadsheet formatting for humans,
      console.log("debug: null price is true, skipping")
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

      //     // break;
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

        var origin = getSelectedValueByName(data, columns, "Origin", i);
        var serial = getSelectedValueByName(data, columns, "Serial #", i);
        var scale = getSelectedValueByName(data, columns, "Item Type", i);
        var body_wood = getSelectedValueByName(data, columns, "Body Wood", i);
        var color = getSelectedValueByName(data, columns, "Color", i);
        var fretboard = getSelectedValueByName(data, columns, "Fretboard", i);
        var pickups = getSelectedValueByName(data, columns, "Pick ups", i);
        var includes_case = getSelectedValueByName(data, columns, "Case", i);
        var moded = getSelectedValueByName(data, columns, "Modded", i);

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
        var template_id = DriveApp.getFilesByName("amp_template").next().getId();
        console.log("generating label", item_name, price);
        var copyId = DriveApp.getFileById(template_id).makeCopy(item_name).getId();
        console.log("log: copyID", copyId);

        // Open the temporary document
        var copyDoc = DocumentApp.openById(copyId);
        // Get the document’s body section
        var label = copyDoc.getBody();

        // Replace the placeholders in the template
        var channel = getSelectedValueByName(data, columns, "Channel", i);
        var power_tubes = getSelectedValueByName(data, columns, "Origin", i);
        var speaker = getSelectedValueByName(data, columns, "Serial #", i);
        var fs = getSelectedValueByName(data, columns, "Item Type", i);
        var effects = getSelectedValueByName(data, columns, "Body Wood", i);
        var cover = getSelectedValueByName(data, columns, "Fretboard", i);

        label.replaceText("{CHANNEL}", channel);
        label.replaceText("{POWER_TUBES}", power_tubes);
        label.replaceText("{SPEAKER}", speaker);
        label.replaceText("{FS}", fs);
        label.replaceText("{EFFECTS}", effects);
        label.replaceText("{COVER}", cover);

        break;

      case "P":
        var template_id = DriveApp.getFilesByName("pedal_template").next().getId();
        console.log("generating label", item_name, price);
        var copyId = DriveApp.getFileById(template_id).makeCopy(item_name).getId();
        console.log("log: copyID", copyId);

        // Open the temporary document
        var copyDoc = DocumentApp.openById(copyId);
        // Get the document’s body section
        var label = copyDoc.getBody();

        // Replace the placeholders in the template
        var box = getSelectedValueByName(data, columns, "Box", i);
        var power = getSelectedValueByName(data, columns, "Power", i);
        label.replaceText("{BOX}", box);
        label.replaceText("{POWER}", power);
        break;

    }

    label.replaceText("{MAKE}", make);
    label.replaceText("{MODEL}", model);
    label.replaceText("{PRICE}", price);
    label.replaceText("{CONDITION}", condition);
    label.replaceText("{NOTES}", notes);
    console.log(":))))");

    copyDoc.saveAndClose();
    var pdfid = docToPDF(copyDoc, datetime);
    console.log(`pdfid returned from docToPDF is ${pdfid}`);
    deleteFileByID(copyId);
    pdfids.push(pdfid);
  }
  // last thing is to combine all exported labels into a single label
  mergePDFs(datetime, pdfids);
  ui.showModalDialog(status("Finished label printing, labels have been saved to google drive"), 'Printing Labels');

}

function printAllLabels() {
  // const datetime = Utilities.formatDate(new Date(), "GMT-7", 'MM-dd-yyyy\'T\'aHH:mm:ss\'Z\'');
  const datetime = Utilities.formatDate(new Date(), "GMT-7", 'EEE, MMM d yyyy h:mm:ss a');

  console.log("Starting print all labels");
  // Get the Google Sheet with the data
  var ss = SpreadsheetApp.getActive().getSheetByName("Sheet4"); // sample sheet comes from what the tab is named
  var list1data = ss.getDataRange().getValues();

  // console.log(columnNames);
  const pdfids = [];
  for (i = 1; i < list1data.length; i++) {

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
    console.log(DriveApp.getFilesByName("misc_template").next().getId());

    switch (item_type) {
      case "M":
        var template_id = DriveApp.getFilesByName("misc_template").next().getId();
        console.log("generating label", item_name, price);
        var copyId = DriveApp.getFileById(template_id).makeCopy(item_name).getId();
        console.log("log: copyID", copyId);

        // Open the temporary document
        var copyDoc = DocumentApp.openById(copyId);
        // Get the document’s body section
        var label = copyDoc.getBody();
        // Replace the placeholders in the template

        break;
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
        var template_id = DriveApp.getFilesByName("amp_template").next().getId();
        console.log("generating label", item_name, price);
        var copyId = DriveApp.getFileById(template_id).makeCopy(item_name).getId();
        console.log("log: copyID", copyId);

        // Open the temporary document
        var copyDoc = DocumentApp.openById(copyId);
        // Get the document’s body section
        var label = copyDoc.getBody();

        // Replace the placeholders in the template
        var channel = getValueByName(ss, "Channel", i);
        var power_tubes = getValueByName(ss, "Origin", i);
        var speaker = getValueByName(ss, "Serial #", i);
        var fs = getValueByName(ss, "Item Type", i);
        var effects = getValueByName(ss, "Body Wood", i);
        var cover = getValueByName(ss, "Fretboard", i);

        label.replaceText("{CHANNEL}", channel);
        label.replaceText("{POWER_TUBES}", power_tubes);
        label.replaceText("{SPEAKER}", speaker);
        label.replaceText("{FS}", fs);
        label.replaceText("{EFFECTS}", effects);
        label.replaceText("{COVER}", cover);

        break;

      case "P":
        var template_id = DriveApp.getFilesByName("pedal_template").next().getId();
        console.log("generating label", item_name, price);
        var copyId = DriveApp.getFileById(template_id).makeCopy(item_name).getId();
        console.log("log: copyID", copyId);

        // Open the temporary document
        var copyDoc = DocumentApp.openById(copyId);
        // Get the document’s body section
        var label = copyDoc.getBody();

        // Replace the placeholders in the template
        var box = getValueByName(ss, "Box", i);
        var power = getValueByName(ss, "Power", i);
        label.replaceText("{BOX}", box);
        label.replaceText("{POWER}", power);
        break;

    }

    label.replaceText("{MAKE}", make);
    label.replaceText("{MODEL}", model);
    label.replaceText("{PRICE}", price);
    label.replaceText("{CONDITION}", condition);
    console.log(":))))");

    copyDoc.saveAndClose();
    var pdfid = docToPDF(copyDoc, datetime);
    console.log(`pdfid returned from docToPDF is ${pdfid}`);
    deleteFileByID(copyId);
    pdfids.push(pdfid);
  }
  // last thing is to combine all exported labels into a single label
  mergePDFs(datetime, pdfids);

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
  pdfBlob.setName(docfile.getName() + "_label.pdf");

  // create new PDF file in Google Drive folder
  folder.createFile(pdfBlob);
  var pdfid = folder.getFilesByName(docfile.getName() + "_label.pdf").next().getId();
  console.log(`doctopdf pdfid ${pdfid}`);
  return pdfid;
}

function deleteFileByID(fileId) {
  var file = Drive.Files.get(fileId);
  if (file.mimeType === MimeType.FOLDER) {
    // possibly ask for confirmation before deleting this folder
  }
  Drive.Files.remove(file.id); // "remove" in Apps Script client library, "delete" elsewhere
}

function getSelectedValueByName(subselection, column_headers, colName, row) {
  var data = subselection
  var col = column_headers.indexOf(colName);
  if (col != -1) {
    return data[row - 1][col];
  }
}

function getValueByName(sheet, colName, row) {
  var data = sheet.getActiveRange().getValues();
  var col = data[1].indexOf(colName);
  if (col != -1) {
    return data[row - 1][col];
  }
}

function mergePDFs(datetime, pdfids) {
  console.log("starting PDF merge function");
  console.log(`datetime set inside pdf merge function is ${datetime}`);

  // get Google Drive folder
  var folder_ID = DriveApp.getFoldersByName('COGS label printer').next().getId();
  console.log("drive app folder id is" + folder_ID);
  var parentFolder = DriveApp.getFolderById(folder_ID); //add this line...
  var folder, folders = DriveApp.getFoldersByName("label_exports " + datetime);

  if (folders.hasNext()) {
    folder = folders.next();
  } else {
  }
  console.log(`folder to be saved to is ${folder}`);

  // Create the final merged PDF file in Drive
  _merge(pdfids, folder)
  // console.log("File saved: "+ " in folder: " + folder.getName());

}

async function _merge(ids, folder) {
  const data = ids.map(id => new Uint8Array(DriveApp.getFileById(id).getBlob().getBytes()));

  const cdnjs = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
  eval(UrlFetchApp.fetch(cdnjs).getContentText().replace(/setTimeout\(.*?,.*?(\d*?)\)/g, "Utilities.sleep($1);return t();"));

  const pdfDoc = await PDFLib.PDFDocument.create();

  for (let i = 0; i < data.length; i++) {
    const pdfData = await PDFLib.PDFDocument.load(data[i]);
    const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
    pages.forEach(page => pdfDoc.addPage(page));
  }

  const bytes = await pdfDoc.save();
  folder.createFile(Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, "merged_labels.pdf"));
}
