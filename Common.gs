function onHomepage(e) {
  console.log('onHomePage = ' + JSON.stringify(e));
  return createHomeCard();
}

function applyDeleteOnFile(file, dryrun, delete_ops, include_subfolder, sheet) {
  var matchStr = getMatchStr(delete_ops);
  var fileName = file.getName();
  if (matchStr === null || fileName.includes(matchStr)) {
    if (dryrun) {
      console.log('Deleting file ' + fileName);
      sheet.appendRow(["File", "Delete", "Yes", include_subfolder, fileName])
      return;
    }

    sheet.appendRow(["File", "Delete", "No", include_subfolder, fileName])
    file.setTrashed(true);
  }
}

function applyDeleteOnFolder(folder, dryrun, delete_ops, include_subfolder, sheet) {
  var matchStr = getMatchStr(delete_ops);
  var notEmpty = folder.getFiles().hasNext();
  var folderName = folder.getName();
  if (matchStr === null || folderName.includes(matchStr)) {
    if (delete_ops.delete_empty_folder) {
      if (notEmpty) return;
    }
    if (dryrun) {
      sheet.appendRow(["Folder", "Delete", "Yes", include_subfolder, folderName])
      console.log('Deleting folder ' + folderName);
      return;
    }

    sheet.appendRow(["Folder", "Delete", "No", include_subfolder, folderName])
    folder.setTrashed(true);
  }
}


function getNewName(name, rename_ops) {
  if (rename_ops.method === "rename_partial") {
    return name.replaceAll(rename_ops.search, rename_ops.replace);
  } 
  
  if (rename_ops.method === "rename_full") {
    return rename_ops.fullname;
  } 
  
  if (rename_ops.method === "rename_adding") {
    return ((rename_ops.before === null)? "" : rename_ops.before) + name + ((rename_ops.after === null)? "" : rename_ops.after);
  } 
  
  throw "Unsupported rename ops = " + rename_ops.rename_method;  
}

function applyRenameOnFile(file, dryrun, rename_ops, include_subfolder, sheet) {
  var matchStr = getMatchStr(rename_ops);
  var fileName = file.getName();
  if (matchStr === null || fileName.includes(matchStr)) {
    var new_name = getNewName(fileName, rename_ops);
    if (dryrun) {
      sheet.appendRow(["File", "Rename", "Yes", include_subfolder, fileName])
      console.log('Renaming file ' + fileName + ' into new name = ' + new_name);
      return;
    }

    sheet.appendRow(["File", "Rename", "No", include_subfolder, fileName])
    file.setName(new_name);
  }
}

function applyRenameOnFolder(folder, dryrun, rename_ops, include_subfolder, sheet) {
  var matchStr = getMatchStr(rename_ops);
  var folderName = folder.getName();
  if (matchStr === null || folderName.includes(matchStr)) {
    var new_name = getNewName(folderName, rename_ops);
    if (dryrun) {
      sheet.appendRow(["Folder", "Rename", "Yes", include_subfolder, folderName])
      console.log('Renaming folder ' + folderName + ' into new name = ' + new_name);
      return;
    }

    sheet.appendRow(["Folder", "Rename", "No", include_subfolder, folderName])
    folder.setName(new_name);  
  }
}

function nextIteration(iterationState, operation, entity, include_subfolder, dryrun, delete_ops, rename_ops, sheet) {
  var currentIteration = iterationState[iterationState.length-1];
  if (currentIteration.fileIteratorContinuationToken !== null) {
    var fileIterator = DriveApp.continueFileIterator(currentIteration.fileIteratorContinuationToken);
    if (fileIterator.hasNext()) {
      var file = fileIterator.next();
      if (entity === "file") {
        if (operation === "delete") {
          applyDeleteOnFile(file, dryrun, delete_ops, include_subfolder, sheet);
        } else if (operation === "rename") {
          applyRenameOnFile(file, dryrun, rename_ops, include_subfolder, sheet);
        } else {
          throw "nextIteration: Unsupported operation type for file = " + operation;
        }
      }
      currentIteration.fileIteratorContinuationToken = fileIterator.getContinuationToken();
      iterationState[iterationState.length-1] = currentIteration;
      return iterationState;
    } 
    
    currentIteration.fileIteratorContinuationToken = null;
    iterationState[iterationState.length-1] = currentIteration;
    return iterationState;
  }

  if (currentIteration.folderIteratorContinuationToken !== null) {
    var folderIterator = DriveApp.continueFolderIterator(currentIteration.folderIteratorContinuationToken);
    if (folderIterator.hasNext()) {
      var folder = folderIterator.next();
      if (entity === "folder") {
        if (operation === "delete") {
          applyDeleteOnFolder(folder, dryrun, delete_ops, include_subfolder, sheet);
        } else if (operation === "rename") {
          applyRenameOnFolder(folder, dryrun, rename_ops, include_subfolder, sheet);
        } else {
          throw "nextIteration: Unsupported operation type for folder = " + operation;
        }
      }
      iterationState[iterationState.length-1].folderIteratorContinuationToken = folderIterator.getContinuationToken();
      if (include_subfolder) {
        iterationState.push(makeIterationFromFolder(folder, operation, entity, delete_ops, rename_ops));
      }
      return iterationState;
    } 
    
    iterationState.pop();
    return iterationState;
  }
}

function getMatchStr(ops) {
  return ops.search;
}

function getMimeType(ops) {
  var type = ops.file_type;
  if (type === "spreadsheet") return MimeType.GOOGLE_SHEETS;
  if (type === "doc") return MimeType.GOOGLE_DOCS;
  if (type === "slide") return MimeType.GOOGLE_SLIDES;
  if (type === "form") return MimeType.GOOGLE_FORMS;
  if (type === "sites") return MimeType.GOOGLE_SITES;
  if (type === "drawing") return MimeType.GOOGLE_DRAWINGS;
  if (type === "appscript") return MimeType.GOOGLE_APPS_SCRIPT;
  if (type === "pdf") return MimeType.PDF;
  if (type === "bmp") return MimeType.BMP;
  if (type === "gif") return MimeType.GIF;
  if (type === "jpeg") return MimeType.JPEG;
  if (type === "png") return MimeType.PNG;
  if (type === "svg") return MimeType.SVG;
  if (type === "css") return MimeType.CSS;
  if (type === "csv") return MimeType.CSV;
  if (type === "html") return MimeType.HTML;
  if (type === "js") return MimeType.JAVASCRIPT;
  if (type === "txt") return MimeType.PLAIN_TEXT;
  if (type === "rtf") return MimeType.RTF;
  if (type === "zip") return MimeType.ZIP;
  if (type === "word1") return MimeType.MICROSOFT_WORD_LEGACY;
  if (type === "word2") return MimeType.MICROSOFT_WORD;
  if (type === "excel1") return MimeType.MICROSOFT_EXCEL_LEGACY;
  if (type === "excel2") return MimeType.MICROSOFT_EXCEL;
  if (type === "ppt1") return MimeType.MICROSOFT_POWERPOINT_LEGACY;
  if (type === "ppt2") return MimeType.MICROSOFT_POWERPOINT;
  if (type === "odt") return MimeType.OPENDOCUMENT_TEXT;
  if (type === "ods") return MimeType.OPENDOCUMENT_SPREADSHEET;
  if (type === "odp") return MimeType.OPENDOCUMENT_PRESENTATION;
  if (type === "odg") return MimeType.OPENDOCUMENT_GRAPHICS;
  return null;
}

function getSearchToken(folder, entity, mimeType) {
    var searchStr = mimeType === null? null : "mimeType='" + mimeType + "'";
    if (searchStr === null) {
      return null;
    }    
    console.log('Get token via search string = ' + searchStr);
    return (entity === "file")? folder.searchFiles(searchStr).getContinuationToken() : folder.searchFolders(searchStr).getContinuationToken();
}

function makeIterationFromFolder(folder, operation, entity, delete_ops, rename_ops) {
  var iteration = {
    folderName: folder.getName(), 
    fileIteratorContinuationToken: null,
    folderIteratorContinuationToken: null
  };

  if (operation === "delete") {
    if (entity === "file") {
      console.log('mime type = ' + getMimeType(delete_ops) + ', match str = ' + getMatchStr(delete_ops));
      var token = getSearchToken(folder, entity, getMimeType(delete_ops));
      if (token !== null) {
        console.log('search file token is not null = ' + token);
        iteration.fileIteratorContinuationToken = token;
      }
    } else {
      var token = getSearchToken(folder, entity, null);
      if (token !== null) {
        iteration.folderIteratorContinuationToken = token;
      }
    }
  } else if (operation === "rename") {
    if (entity === "file") {
      var token = getSearchToken(folder, entity, getMimeType(rename_ops));
      if (token !== null) {
        iteration.fileIteratorContinuationToken = token;
      }
    } else {
      var token = getSearchToken(folder, entity, null);
      if (token !== null) {
        iteration.folderIteratorContinuationToken = token;
      }
    }
  }

  if (iteration.fileIteratorContinuationToken === null) {
    iteration.fileIteratorContinuationToken = folder.getFiles().getContinuationToken();
  }

  if (iteration.folderIteratorContinuationToken === null) {
    iteration.folderIteratorContinuationToken = folder.getFolders().getContinuationToken();
  }

  /// now iteration has both file token and folder token
  /// when entity is folder only do not set file token because it is useless
  if (entity === "folder") {
    iteration.fileIteratorContinuationToken = null;
  }

  return iteration;
}

function getJobStatusKey(email) {
  return email + ":job";
}

function getJobMetadataKey(email) {
  return email + ":metadata";
}

function findTrigger(id) {
  console.log('findTrigger: finding trigger with id = ' + id);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getUniqueId().toString() === id) {
      return triggers[i];
    }
  }
  return null;
}

function onScheduledRun(e) {
  var triggerId = e.triggerUid.toString();
  var email = Session.getEffectiveUser().getEmail();
  var jobMetadataKey = getJobMetadataKey(email);
  var properties = PropertiesService.getUserProperties();
  var metadata = properties.getProperty(jobMetadataKey);
  if (metadata === null) {
    console.log('onScheduleRun: Cannot find job metadata for key ' + jobMetadataKey + '. Skip since no work to do');
    return;
  }

  var json = JSON.parse(metadata);
  var folder = DriveApp.getFolderById(json.folder.id);
  var spreadsheet = SpreadsheetApp.openById(json.report_spreadsheet.id);
  var sheet = spreadsheet.getSheetByName(getOperationLogSheetName());
  var progress = spreadsheet.getSheetByName()
  iterateFolder(folder, json.operation, json.entity, json.include_subfolder, json.dry_run, json.delete_ops, json.rename_ops, sheet, progress);
  var trigger = findTrigger(triggerId);
  if (trigger !== null) {
    ScriptApp.deleteTrigger(trigger);
  }
}

function iterateFolder(folder, operation, entity, include_subfolder, dryrun, delete_ops, rename_ops, sheet, progress) {
  var email = Session.getEffectiveUser().getEmail();
  console.log('Iterate entity in folder: email = ' + email + ', folder = ' + folder + ', operation = ' + operation + ', entity = ' + entity + ', include_subfolder = ' + include_subfolder + ', dryrun = ' + dryrun + ', delete_ops = ' + JSON.stringify(delete_ops) + ', rename_ops = ' + JSON.stringify(rename_ops));
  var MAX_RUNNING_TIME_MS = 4.5 * 60 * 1000;
  var startTime = (new Date()).getTime();
  var jobKey = getJobStatusKey(email);
  var properties = PropertiesService.getUserProperties();
  var iterationState = JSON.parse(properties.getProperty(jobKey));
  if (iterationState !== null) {
    if (folder.getName() !== iterationState[0].folderName) {
      console.error("Iterating a new folder: " + folder.getName() + ". End early since existing operation is not done.");
      return;
    }
    console.info("Resuming iteration for folder: " + folder.getName());
  }
  if (iterationState === null) {
    console.info("Starting new iteration for folder: " + folder.getName());
    progress.appendRow([(new Date()).getTime(), "STARTED"]);
    iterationState = [];
    iterationState.push(makeIterationFromFolder(folder, operation, entity, delete_ops, rename_ops));
  }  

  progress.appendRow([(new Date()).getTime(), "STARTED NEW ITERATION"]);
  while (iterationState.length > 0) {
    iterationState = nextIteration(iterationState, operation, entity, include_subfolder, dryrun, delete_ops, rename_ops, sheet);
    var currTime = (new Date()).getTime();
    var elapsedTimeInMS = currTime - startTime;
    var timeLimitExceeded = elapsedTimeInMS >= MAX_RUNNING_TIME_MS;
    if (timeLimitExceeded) {
      properties.setProperty(jobKey, JSON.stringify(iterationState));
      console.info("Stopping loop after '%d' milliseconds.", elapsedTimeInMS);
      progress.appendRow([(new Date()).getTime(), "ENDED NEW ITERATION"]);
      // Continue work after 1s
      ScriptApp.newTrigger("onScheduledRun").timeBased().after(1000).create();
      return;
    }
  }

  console.info("Done iterating. Deleting iterating state ... ");
  progress.appendRow([(new Date()).getTime(), "DONE"]);
  properties.deleteProperty(jobKey);
  properties.deleteProperty(jobMetadataKey);
}


function parseFolderFromEvent(e) {
  var folderId = null;
  if (('selectedItems' in e.drive) && (e.drive.activeCursorItem.mimeType === 'application/vnd.google-apps.folder')) {
    folderId = e.drive.activeCursorItem.id;
  }
  console.log('folderId = ' + folderId);
  if (folderId === null) {
    var folder = DriveApp.getRootFolder();
  } else {
    var folder = DriveApp.getFolderById(folderId);
  }
  return folder;
}

function getOperationLogSheetName() {
  return "Operation Logs";
}

function getProgressSheetName() {
  return "Progress";
}

function deleteFileHandler(e) {
  console.log('deleteFileHandler = ' + JSON.stringify(e));
  var email = Session.getEffectiveUser().getEmail();
  var folder = parseFolderFromEvent(e);
  var filename_match = ('file_name_field' in e.formInput)? e.formInput.file_name_field : null;
  var delete_ops = {
    file_type: e.formInput.file_type_field,
    search: filename_match,
    delete_empty_folder: null,
  }
  var dryrun = JSON.parse(e.parameters.dryrun);
  var include_subfolder = JSON.parse(e.parameters.include_subfolder);
  var properties = PropertiesService.getUserProperties();
  var jobMetadataKey = getJobMetadataKey(email);
  var spreadsheet = SpreadsheetApp.create("DriveWorks_progress_report_" + Date.now());
  var sheet = spreadsheet.insertSheet(getOperationLogSheetName());
  sheet.appendRow(["Entity", "Operation", "Dry run", "Include subfolder", "Entity name"]);
  var progress = spreadsheet.insertSheet(getProgressSheetName());
  progress.appendRow(["Time", "Progress"]);
  var metadata = {
      folder: {
        id: folder.getId(),
        name: folder.getName()
      },
      operation : "delete", 
      entity : "file",
      include_subfolder : include_subfolder,
      dry_run : dryrun,
      delete_ops : delete_ops,
      rename_ops : null,
      report_spreadsheet : {
        id : spreadsheet.getId(),
        name: spreadsheet.getName()
      }
  };
  if (properties.getProperty(jobMetadataKey) === null) {
    properties.setProperty(jobMetadataKey, JSON.stringify(metadata));
  } else {
    console.log('job metadata key exists, but it should not, the key =' + jobMetadataKey);
  }
  ScriptApp.newTrigger("onScheduledRun").timeBased().after(1000).create();
  var card = createDeleteFileCard(include_subfolder, dryrun);
  var navigation = CardService.newNavigation().updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

function deleteJob() {
  var email = Session.getEffectiveUser().getEmail();
  var jobMetadataKey = getJobMetadataKey(email);
  var properties = PropertiesService.getUserProperties();
  properties.deleteProperty(jobMetadataKey);
  var navigation = CardService.newNavigation()
      .popCard();
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}


function buildCardViaPropertiesIfExist() {
  var properties = PropertiesService.getUserProperties();
  var email = Session.getEffectiveUser().getEmail();
  var jobMetadataKey = getJobMetadataKey(email);
  var metadataStr = properties.getProperty(jobMetadataKey);
  if (metadataStr === null) {
    return null;
  }
  var metadata = JSON.parse(metadataStr);
  var cardHeader = CardService.newCardHeader()  
    .setTitle("DriveWorks")
    .setSubtitle("Job Settings and Status")
    .setImageUrl(getLogoURL());
  var paymentSection = getPaymentSection();
  var mainSection = CardService.newCardSection(); 
  var operationType = CardService.newDecoratedText().setText("Drive operation type: <b><font color='#065fd4'>" + metadata.operation + "</b>").setWrapText(true).setTopLabel("Primary Job Settings");
  var entityType = CardService.newDecoratedText().setText("Entity type: <b><font color='#065fd4'>" + metadata.entity + "</b>").setWrapText(true);
  var folderName = CardService.newDecoratedText().setText("Target folder: <b><font color='#065fd4'>" + metadata.folder.name + "</b>");
  var dryrun = CardService.newDecoratedText().setText("Is dryrun: <b><font color='#065fd4'>" + metadata.dry_run + "</b>");
  var include_subfolder = CardService.newDecoratedText().setText("Include subfolder: <b><font color='#065fd4'>" + metadata.include_subfolder + "</b>");
  mainSection
    .addWidget(operationType)
    .addWidget(entityType)
    .addWidget(folderName)
    .addWidget(dryrun)
    .addWidget(include_subfolder);
  var spreadsheet = SpreadsheetApp.openById(metadata.report_spreadsheet.id);
  var link = CardService.newOpenLink()
        .setUrl(spreadsheet.getUrl())
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.RELOAD_ADD_ON)
  var status = CardService.newDecoratedText().setText("<b><font color='#065fd4'>CLICK HERE</b> to monitor job progress. You will be notified via email once the job completes. You are not allowed to start a new job while a previous job is running unless the previous is deleted.").setWrapText(true).setTopLabel("Existing Job Status").setOpenLink(link);
  mainSection.addWidget(status);  
  var action = CardService.newAction()
    .setFunctionName('deleteJob')
  var button = CardService.newTextButton()
    .setText('Delete Job')
    .setOnClickAction(action)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  var buttonSet = CardService.newButtonSet()
    .addButton(button);
  mainSection.addWidget(buttonSet);
  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(paymentSection)
    .addSection(mainSection);

  return card; 
}

function deleteFolderHandler(e) {
  console.log('deleteFolderHandler = ' + JSON.stringify(e));
  var folder = parseFolderFromEvent(e);
  var foldername_match = ('folder_name_field' in e.formInput)? e.formInput.folder_name_field : null;
  var delete_empty_folder = ('delete_empty_folders_field' in e.formInput)? true : false;
  var delete_ops = {
    file_type: null,
    search: foldername_match,
    delete_empty_folder: delete_empty_folder,
  }
  var email = Session.getEffectiveUser().getEmail();
  var dryrun = JSON.parse(e.parameters.dryrun);
  var include_subfolder = JSON.parse(e.parameters.include_subfolder);
  var properties = PropertiesService.getUserProperties();
  var jobMetadataKey = getJobMetadataKey(email);
  var spreadsheet = SpreadsheetApp.create("DriveWorks_progress_report_" + Date.now());
  var sheet = spreadsheet.insertSheet(getOperationLogSheetName());
  sheet.appendRow(["Entity", "Operation", "Dry run", "Include subfolder", "Entity name"]);
  var progress = spreadsheet.insertSheet(getProgressSheetName());
  progress.appendRow(["Time", "Progress"]);
  var metadata = {
      folder: {
        id: folder.getId(),
        name: folder.getName()
      },
      operation : "delete", 
      entity : "folder",
      include_subfolder : include_subfolder,
      dry_run : dryrun,
      delete_ops : delete_ops,
      rename_ops : null,
      report_spreadsheet : {
        id : spreadsheet.getId(),
        name: spreadsheet.getName()
      }
  };
  if (properties.getProperty(jobMetadataKey) === null) {
    properties.setProperty(jobMetadataKey, JSON.stringify(metadata));
  } else {
    console.log('job metadata key exists, but it should not, the key =' + jobMetadataKey);
  }
  ScriptApp.newTrigger("onScheduledRun").timeBased().after(1000).create();
  var card = createDeleteFolderCard(include_subfolder, dryrun);
  var navigation = CardService.newNavigation().updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

function renameFileHandler(e) {
  console.log('renameFileHandler = ' + JSON.stringify(e));  
  var folder = parseFolderFromEvent(e);
  var rename_ops = {
    method: e.formInput.rename_method_field,  
    file_type: e.formInput.file_type_field,
    search: ("file_name_search_field" in e.formInput)? e.formInput.file_name_search_field : null,
    replace: ("file_name_replace_field" in e.formInput)? e.formInput.file_name_replace_field : null,
    fullname: ("new_file_name_field" in e.formInput)? e.formInput.new_file_name_field : null,
    before: ("file_name_before_field" in e.formInput)? e.formInput.file_name_before_field: null,
    after: ("file_name_after_field" in e.formInput)? e.formInput.file_name_after_field : null,
  }
  
  var email = Session.getEffectiveUser().getEmail();
  var dryrun = JSON.parse(e.parameters.dryrun);
  var include_subfolder = JSON.parse(e.parameters.include_subfolder);
  var properties = PropertiesService.getUserProperties();
  var jobMetadataKey = getJobMetadataKey(email);
  var spreadsheet = SpreadsheetApp.create("DriveWorks_progress_report_" + Date.now());
  var sheet = spreadsheet.insertSheet(getOperationLogSheetName());
  sheet.appendRow(["Entity", "Operation", "Dry run", "Include subfolder", "Entity name"]);
  var progress = spreadsheet.insertSheet(getProgressSheetName());
  progress.appendRow(["Time", "Progress"]);
  var metadata = {
      folder: {
        id: folder.getId(),
        name: folder.getName()
      },
      operation : "rename", 
      entity : "file",
      include_subfolder : include_subfolder,
      dry_run : dryrun,
      delete_ops : null,
      rename_ops : rename_ops,
      report_spreadsheet : {
        id : spreadsheet.getId(),
        name: spreadsheet.getName()
      }
  };
  if (properties.getProperty(jobMetadataKey) === null) {
    properties.setProperty(jobMetadataKey, JSON.stringify(metadata));
  } else {
    console.log('job metadata key exists, but it should not, the key =' + jobMetadataKey);
  }
  ScriptApp.newTrigger("onScheduledRun").timeBased().after(1000).create();
  var card = createRenameFileCard(e.formInput.rename_method_field, include_subfolder, dryrun);
  var navigation = CardService.newNavigation().updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

function renameFolderHandler(e) {
  console.log('renameFolderHandler = ' + JSON.stringify(e));
  var folder = parseFolderFromEvent(e);
  var rename_ops = {
    method: e.formInput.rename_method_field,  
    file_type: null,
    search: ("folder_name_search_field" in e.formInput)? e.formInput.folder_name_search_field : null,
    replace: ("folder_name_replace_field" in e.formInput)? e.formInput.folder_name_replace_field : null,
    fullname: ("new_folder_name_field" in e.formInput)? e.formInput.new_folder_name_field : null,
    before: ("folder_name_before_field" in e.formInput)? e.formInput.folder_name_before_field: null,
    after: ("folder_name_after_field" in e.formInput)? e.formInput.folder_name_after_field : null,
  }
  var email = Session.getEffectiveUser().getEmail();
  var dryrun = JSON.parse(e.parameters.dryrun);
  var include_subfolder = JSON.parse(e.parameters.include_subfolder);
  var properties = PropertiesService.getUserProperties();
  var jobMetadataKey = getJobMetadataKey(email);
  var spreadsheet = SpreadsheetApp.create("DriveWorks_progress_report_" + Date.now());
  var sheet = spreadsheet.insertSheet(getOperationLogSheetName());
  sheet.appendRow(["Entity", "Operation", "Dry run", "Include subfolder", "Entity name"]);
  var progress = spreadsheet.insertSheet(getProgressSheetName());
  progress.appendRow(["Time", "Progress"]);
  spreadsheet.deleteSheet(spreadsheet.getSheetByName('Sheet1'));
  var metadata = {
      folder: {
        id: folder.getId(),
        name: folder.getName()
      },
      operation : "rename", 
      entity : "folder",
      include_subfolder : include_subfolder,
      dry_run : dryrun,
      delete_ops : null,
      rename_ops : rename_ops,
      report_spreadsheet : {
        id : spreadsheet.getId(),
        name: spreadsheet.getName()
      }
  };
  if (properties.getProperty(jobMetadataKey) === null) {
    properties.setProperty(jobMetadataKey, JSON.stringify(metadata));
  } else {
    console.log('job metadata key exists, but it should not, the key =' + jobMetadataKey);
  }
  ScriptApp.newTrigger("onScheduledRun").timeBased().after(1000).create();
  var card = createRenameFolderCard(e.formInput.rename_method_field, include_subfolder, dryrun);
  var navigation = CardService.newNavigation().updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

function getFileTypeWidget() {
  return CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Delete when file type matches")
    .setFieldName("file_type_field")
    .addItem("Any file type", "all", true)
    .addItem("Google Sheets", "spreadsheet", false)
    .addItem("Google Docs", "doc", false)
    .addItem("Google Slides", "slide", false)
    .addItem("Google Forms", "form", false)
    .addItem("Google Sites", "sites", false)
    .addItem("Google Drawings", "drawing", false)
    .addItem("Google App Script", "appscript", false)
    .addItem("Adobe PDF (.pdf)", "pdf", false)
    .addItem("BMP file (.bmp)", "bmp", false)
    .addItem("GIF file (.gif)", "gif", false)
    .addItem("JPEG file (.jpeg)", "jpeg", false)
    .addItem("PNG file (.png)", "png", false)
    .addItem("SVG file (.svg)", "svg", false)
    .addItem("CSS file (.css)", "css", false)
    .addItem("CSV file (.csv)", "csv", false)
    .addItem("Html file (.html)", "html", false)
    .addItem("Javascript file (.js)", "js", false)
    .addItem("Plain Text (.txt)", "txt", false)
    .addItem("Rich Text file (.rtf)", "rtf", false)
    .addItem("ZIP file (.zip)", "zip", false)
    .addItem("Microsoft Word (.doc)", "word1", false)
    .addItem("Microsoft Word (.docx)", "word2", false)
    .addItem("Microsoft Excel (.xls)", "excel1", false)
    .addItem("Microsoft Excel (.xlsx)", "excel2", false)
    .addItem("Microsoft Powerpoint (.ppt)", "ppt1", false)
    .addItem("Microsoft Powerpoint (.pptx)", "ppt2", false)
    .addItem("OpenDocument Text (.odt)", "odt", false)
    .addItem("OpenDocument Spreadsheet (.ods)", "ods", false)
    .addItem("OpenDocument Presentation (.odp)", "odp", false)
    .addItem("OpenDocument Graphics (.odg)", "odg", false);
}

function createDeleteFileCard(include_subfolder, dryrun) {
  var statusCard = buildCardViaPropertiesIfExist();
  if (statusCard) {
    return statusCard.build();
  }
  var cardHeader = CardService.newCardHeader()  
    .setTitle("DriveWorks")
    .setSubtitle("Delete files")
    .setImageUrl(getLogoURL());
  var paymentSection = getPaymentSection();
  var mainSection = CardService.newCardSection().setHeader("Set filters for files to delete");

  var filenameMatch = CardService.newTextInput()
    .setFieldName("file_name_field")
    .setTitle("Delete when file name contains input");
  mainSection.addWidget(filenameMatch);

  var fileType = getFileTypeWidget();
  mainSection.addWidget(fileType);

  var action = CardService.newAction()
    .setFunctionName('deleteFileHandler')
    .setParameters({include_subfolder: JSON.stringify(include_subfolder), dryrun : JSON.stringify(dryrun)});
  var button = CardService.newTextButton()
    .setText('Delete Files')
    .setOnClickAction(action)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  var buttonSet = CardService.newButtonSet()
    .addButton(button);
  mainSection.addWidget(buttonSet);
  var status = CardService.newDecoratedText().setText("File deletion progress is going to be shown in a spreadsheet once button is clicked. You will be notified via email once the job completes.").setWrapText(true);  
  mainSection.addWidget(status);

  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(paymentSection)
    .addSection(mainSection);

  return card.build();  
}

function createDeleteFolderCard(include_subfolder, dryrun) {
  var statusCard = buildCardViaPropertiesIfExist();
  if (statusCard) {
    return statusCard.build();
  }
  var cardHeader = CardService.newCardHeader()  
    .setTitle("DriveWorks")
    .setSubtitle("Delete folders")
    .setImageUrl(getLogoURL());
  var paymentSection = getPaymentSection();
  var mainSection = CardService.newCardSection().setHeader("Set filters for folders to delete");

  var deleteEmpty = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("delete_empty_folders_field")
    .addItem("Delete empty folders only", "delete_empty_folder", true);
  mainSection.addWidget(deleteEmpty);

  var foldernameMatch = CardService.newTextInput()
    .setFieldName("folder_name_field")
    .setTitle("Delete when folder name contains input");
  mainSection.addWidget(foldernameMatch);

  var action = CardService.newAction()
    .setFunctionName('deleteFolderHandler')
    .setParameters({include_subfolder: JSON.stringify(include_subfolder), dryrun : JSON.stringify(dryrun)});

  var button = CardService.newTextButton()
    .setText('Delete Folders')
    .setOnClickAction(action)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  var buttonSet = CardService.newButtonSet()
    .addButton(button);
  mainSection.addWidget(buttonSet);

  var status = CardService.newDecoratedText().setText("Folder deletion progress is going to be shown in a spreadsheet once button is clicked. You will be notified via email once the job completes.").setWrapText(true).setTopLabel("Folder deletion status is shown below");
  mainSection.addWidget(status);

  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(paymentSection)
    .addSection(mainSection);

  return card.build();  
}

function changeFileRenameHandler(e) {
  console.log('changeFileRenameHandler = ' + JSON.stringify(e));
  var card = createRenameFileCard(e.formInput.rename_method_field, JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun));
  var navigation = CardService.newNavigation().updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

function createRenameFileCard(rename_method="rename_partial", include_subfolder, dryrun) {
  var statusCard = buildCardViaPropertiesIfExist();
  if (statusCard) {
    return statusCard.build();
  }
  var cardHeader = CardService.newCardHeader()  
    .setTitle("DriveWorks")
    .setSubtitle("Rename file")
    .setImageUrl(getLogoURL());
  var paymentSection = getPaymentSection();
  var mainSection = CardService.newCardSection().setHeader("Set filters to rename files");
  var action = CardService.newAction()
    .setFunctionName('changeFileRenameHandler')
    .setParameters({include_subfolder: JSON.stringify(include_subfolder), dryrun : JSON.stringify(dryrun)});
  var renameMethod = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setFieldName("rename_method_field")
    .setOnChangeAction(action);
  if (rename_method === "rename_partial") {
    renameMethod
     .addItem("Replace matched file name", "rename_partial", true)
     .addItem("Rename full name", "rename_full", false)
     .addItem("Add string before or after file name", "rename_adding", false);
  } else if (rename_method === "rename_full") {
    renameMethod
     .addItem("Replace matched file name", "rename_partial", false)
     .addItem("Rename full name", "rename_full", true)
     .addItem("Add string before or after file name", "rename_adding", false);    
  } else if (rename_method === "rename_adding") {
    renameMethod
     .addItem("Replace matched file name", "rename_partial", false)
     .addItem("Rename full name", "rename_full", false)
     .addItem("Add string before or after file name", "rename_adding", true);
  } else {
    throw "Unsupported rename method = " + rename_method;
  }

  mainSection.addWidget(renameMethod);  
  if (rename_method === "rename_partial") {
    var search = CardService.newTextInput()
    .setFieldName("file_name_search_field")
    .setTitle("String to match file name");
    var replace = CardService.newTextInput()
    .setFieldName("file_name_replace_field")
    .setTitle("String to replace the matches");
    mainSection.addWidget(search).addWidget(replace);    
  } else if (rename_method === "rename_full") {
    var newname = CardService.newTextInput()
    .setFieldName("new_file_name_field")
    .setTitle("New file name");
    mainSection.addWidget(newname);
  } else if (rename_method === "rename_adding") {
    var before = CardService.newTextInput()
    .setFieldName("file_name_before_field")
    .setTitle("Add string before file name");
    var after = CardService.newTextInput()
    .setFieldName("file_name_after_field")
    .setTitle("Add string after file name");
    mainSection.addWidget(before).addWidget(after);
  } else {
    throw "Unsupported rename method X = " + rename_method;
  }

  var fileType = getFileTypeWidget();
  mainSection.addWidget(fileType);

  var action = CardService.newAction()
    .setFunctionName('renameFileHandler')
    .setParameters({include_subfolder: JSON.stringify(include_subfolder), dryrun : JSON.stringify(dryrun)});
  var button = CardService.newTextButton()
    .setText('Rename Files')
    .setOnClickAction(action)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  var buttonSet = CardService.newButtonSet()
    .addButton(button);
  mainSection.addWidget(buttonSet);
  var status = CardService.newDecoratedText().setText("File renaming progress is going to be shown in a spreadsheet once button is clicked. You will be notified via email once the job completes.").setWrapText(true).setTopLabel("File renaming status is shown below");
  mainSection.addWidget(status);
  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(paymentSection)
    .addSection(mainSection);

  return card.build();  
}

function changeFolderRenameHandler(e) {
  console.log('changeFolderRenameHandler = ' + JSON.stringify(e));
  var card = createRenameFolderCard(e.formInput.rename_method_field, JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun));
  var navigation = CardService.newNavigation().updateCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

function createRenameFolderCard(rename_method="rename_partial", include_subfolder, dryrun) {
  var statusCard = buildCardViaPropertiesIfExist();
  if (statusCard) {
    return statusCard.build();
  }
  var cardHeader = CardService.newCardHeader()  
    .setTitle("DriveWorks")
    .setSubtitle("Rename folder")
    .setImageUrl(getLogoURL());
  var paymentSection = getPaymentSection();
  var mainSection = CardService.newCardSection().setHeader("Set filters to rename folders");
  var action = CardService.newAction()
    .setFunctionName('changeFolderRenameHandler')
    .setParameters({include_subfolder: JSON.stringify(include_subfolder), dryrun: JSON.stringify(dryrun)});
  var renameMethod = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setFieldName("rename_method_field")
    .setOnChangeAction(action);
  if (rename_method === "rename_partial") {
    renameMethod
     .addItem("Replace matched folder name", "rename_partial", true)
     .addItem("Rename full name", "rename_full", false)
     .addItem("Add string before or after folder name", "rename_adding", false);
  } else if (rename_method === "rename_full") {
    renameMethod
     .addItem("Replace matched folder name", "rename_partial", false)
     .addItem("Rename full name", "rename_full", true)
     .addItem("Add string before or after folder name", "rename_adding", false);    
  } else if (rename_method === "rename_adding") {
    renameMethod
     .addItem("Replace matched folder name", "rename_partial", false)
     .addItem("Rename full name", "rename_full", false)
     .addItem("Add string before or after folder name", "rename_adding", true);
  } else {
    throw "Unsupported rename method = " + rename_method;
  }

  mainSection.addWidget(renameMethod);  
  if (rename_method === "rename_partial") {
    var search = CardService.newTextInput()
    .setFieldName("folder_name_search_field")
    .setTitle("String to match folder name");
    var replace = CardService.newTextInput()
    .setFieldName("folder_name_replace_field")
    .setTitle("String to replace the matches");
    mainSection.addWidget(search).addWidget(replace);    
  } else if (rename_method === "rename_full") {
    var newname = CardService.newTextInput()
    .setFieldName("new_folder_name_field")
    .setTitle("New folder name");
    mainSection.addWidget(newname);
  } else if (rename_method === "rename_adding") {
    var before = CardService.newTextInput()
    .setFieldName("folder_name_before_field")
    .setTitle("Add string before folder name");
    var after = CardService.newTextInput()
    .setFieldName("folder_name_after_field")
    .setTitle("Add string after folder name");
    mainSection.addWidget(before).addWidget(after);
  } else {
    throw "Unsupported rename method X = " + rename_method;
  }

  var action = CardService.newAction()
    .setFunctionName('renameFolderHandler')
    .setParameters({include_subfolder: JSON.stringify(include_subfolder), dryrun : JSON.stringify(dryrun)});

  var button = CardService.newTextButton()
    .setText('Rename Folders')
    .setOnClickAction(action)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  var buttonSet = CardService.newButtonSet()
    .addButton(button);
  mainSection.addWidget(buttonSet);
  var status = CardService.newDecoratedText().setText("Folder renaming progress is going to be shown in a spreadsheet once button is clicked. You will be notified via email once the job completes.").setWrapText(true).setTopLabel("Folder renaming status is shown below");
  mainSection.addWidget(status);
  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(paymentSection)
    .addSection(mainSection);

  return card.build();  
}

function configureMore(e) {
  console.log('configureMore = ' + JSON.stringify(e));
  var expire_time = getExpireTime();
  if (expire_time!== -1 && expire_time <= Date.now()/1000) {
    var card = getBuymoreCard();
    var navigation = CardService.newNavigation()
      .pushCard(card);
    var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
    return actionResponse.build();
  }

  var operation_type = e.formInput.drive_operation_type_field;
  var entity_type = e.formInput.entity_type_field;
  var include_subfolder = e.formInput.include_subfolders_field === "include_subfolder";
  var dryrun = e.formInput.dryrun_field === "dryrun";
  var card = null;
  if (operation_type === "delete") {
    var card = (entity_type === "file")? createDeleteFileCard(include_subfolder, dryrun) : createDeleteFolderCard(include_subfolder, dryrun);
  } 
  
  if (operation_type === "rename") {
    var card = (entity_type === "file")? createRenameFileCard("rename_partial", include_subfolder, dryrun) : createRenameFolderCard("rename_partial", include_subfolder, dryrun);    
  }

  if (card === null) {
    throw "Unsupported operation type = " + operation_type;
  }
 
  var navigation = CardService.newNavigation()
      .pushCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

function getExpireTime(version=1) {
  try {
    var initial_days = 7;
    var email = Session.getEffectiveUser().getEmail();
    var key = email + version.toString();
    var properties = PropertiesService.getScriptProperties();
    var expire_time = properties.getProperty(key);
    if (expire_time === null) {
      var now = Date.now()/1000;
      var future = now + initial_days*24*3600; 
      expire_time = properties.setProperty(key, future.toString()).getProperty(key);
    } 
    console.log('expire_time = ' + expire_time);
    var ret = parseInt(expire_time);
    console.log('getExpireTime:' + key + ', initial_days = ' + initial_days + ', expire_time = ' + expire_time);
    return ret;
  } catch (e) {
    console.log('getExpireTime error:' + e);
    throw e;
  }
}

function getRemainingDays() {
  try {
    var email = Session.getEffectiveUser().getEmail(); 
    var expire_time = getExpireTime();
    if (expire_time === -1) {
      return -1;
    }
    var now = Date.now()/1000;
    if (expire_time <= now) {
      return 0;
    }

    var days = Math.ceil((expire_time - now) / (24*3600));
    console.log('Email:' + email + ',remaining days = ' + days);
    return days;
  } catch (e) {
    console.log('getRemainingDays error:' + e);
    throw e;
  }
}

function getBuymoreCard() {
  var email = Session.getEffectiveUser().getEmail(); 
  var cardHeader = CardService.newCardHeader()  
    .setTitle("DriveWorks")
    .setSubtitle("Buy More")
    .setImageUrl(getLogoURL());
  var mainSection = CardService.newCardSection().setHeader("Buy more days for app access");
  var description = CardService.newDecoratedText().setText("To support users and add features, we are kindly asking to pay for its use. We hope you find the subscription price reasonable.").setWrapText(true);
  mainSection.addWidget(description);

  var options = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setTitle("Select payment option")
    .setFieldName("payment_option_field")
    .addItem("$5 per month (billed per month)", "month", false)
    .addItem("$50 per year (billed per year)", "year", true);
  mainSection.addWidget(options);

  var paypalLink = CardService.newOpenLink()
        .setUrl("https://www.paypal.com/paypalme/SmartGAS888")
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.RELOAD_ADD_ON);
  var paypal = CardService.newDecoratedText().setText("Option1 (PayPal): Please use <b><font color='#065fd4'>PayPal</font></b> and send an email to james.cui.code@gmail.com after the payment to update your access.").setWrapText(true).setOpenLink(paypalLink);
  mainSection.addWidget(paypal);

  var buycoffeelink = CardService.newOpenLink()
        .setUrl("https://www.buymeacoffee.com/smartgas")
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.RELOAD_ADD_ON)
  var buycoffee = CardService.newDecoratedText().setText("Option2 (Buy Me A Coffee): Please use <b><font color='#065fd4'>Buy me a coffee</font></b> and send an email to james.cui.code@gmail.com after the payment to update your access.").setWrapText(true).setOpenLink(buycoffeelink);
  mainSection.addWidget(buycoffee);

  if (email === "james.cui.code@gmail.com") {
    var email_account = CardService.newTextInput()
      .setFieldName("email_field")
      .setTitle("Email account");
    var added_days = CardService.newTextInput()
      .setFieldName("added_days")
      .setTitle("Days");
    mainSection.addWidget(email_account).addWidget(added_days);
    var setDaysAction = CardService.newAction().setMethodName("setDays")
    var button = CardService.newTextButton()
      .setText('Set Days')
      .setOnClickAction(setDaysAction)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
    var buttonSet = CardService.newButtonSet()
      .addButton(button);
    mainSection.addWidget(buttonSet);
  }

  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(mainSection);

  return card.build();  
}

/*
formInput: 
   { added_days: '10',
     email_field: 'james.cui.code@gmail.com',
     payment_option_field: 'year' },
*/

function setDays(e) {
  var email = e.formInput.email_field;
  var days = parseInt(e.formInput.added_days);
  setExpiredTime(email, 1, days);
}

function setExpiredTime(user, version, days) {
  try {
    var key = user + version.toString();
    var properties = PropertiesService.getScriptProperties();
    if (days === -1) {
      properties.setProperty(key, "-1");
      return -1;
    }
    var expire_time = properties.getProperty(key);
    var now = Date.now()/1000;
    var future = now + days*24*3600; 
    expire_time = properties.setProperty(key, future.toString()).getProperty(key);
    var ret = parseInt(expire_time);
    console.log('setExpiredTime: key = ' + key + ', expire_time = ' + expire_time);
    return ret;
  } catch (e) {
    console.log('setExpiredTime error:' + e);
    return -1;
  }
}

function buymore() {
  var card = getBuymoreCard();
  var navigation = CardService.newNavigation()
      .pushCard(card);
  var actionResponse = CardService.newActionResponseBuilder()
      .setNavigation(navigation);
  return actionResponse.build();
}

function getPaymentSection() {
  var endIcon = CardService.newIconImage().setIcon(CardService.Icon.VIDEO_PLAY);
  var buymoreIcon = CardService.newIconImage().setIcon(CardService.Icon.DOLLAR);
  var remainDays = getRemainingDays();
  var action = CardService.newAction().setFunctionName('buymore');
  var buymore = CardService.newDecoratedText().setText("Remaining <b><font color='#065fd4'>" + remainDays + " days</font></b> for trial").setStartIcon(buymoreIcon).setEndIcon(endIcon).setWrapText(true).setOnClickAction(action);
  return CardService.newCardSection().addWidget(buymore);  
}

function getLogoURL() {
  return "https://lh3.googleusercontent.com/drive-viewer/AITFw-yAbBqZJMHtHd8akkZr8Ri1KrpInZoT1671AEt1x4enG_OOOFrP8_-4rvMbb3oU6jzVIBOF1FFRNAWdE3F0V6hxbydf=s2560";
}

function createHomeCard(item={}) {
  var entityType = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setTitle("Select drive entity type")
    .setFieldName("entity_type_field")
    .addItem("File", "file", true)
    .addItem("Folder", "folder", false);
  var operationType = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setTitle("Select drive operation type")
    .setFieldName("drive_operation_type_field")
    .addItem("Delete", "delete", true)
    .addItem("Rename", "rename", false);
  var textStart = "Selected operation folder<br><b><font color='#065fd4'>";
  var textEnd = "</font></b>"; 
  if ('title' in item) {
    var text = textStart + item.title + textEnd;
  } else {
    var text = textStart + "My Drive" + textEnd;
  }
  var selectedFolder = CardService.newDecoratedText().setText(text).setWrapText(true).setBottomLabel("Change by selecting a different folder");

  var dryrun = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("dryrun_field")
    .addItem("Preview operations only without executing (dryrun)", "dryrun", true);

  var includeSubfolder = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("include_subfolders_field")
    .addItem("Apply to subfolders", "include_subfolder", true);

  var configureMoreAction = CardService.newAction()
      .setFunctionName('configureMore');
  var button = CardService.newTextButton()
      .setText('Configure Details')
      .setOnClickAction(configureMoreAction)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  var buttonSet = CardService.newButtonSet()
      .addButton(button);
  var mainSection = CardService.newCardSection()
    .addWidget(operationType)
    .addWidget(entityType)
    .addWidget(selectedFolder)
    .addWidget(dryrun)
    .addWidget(includeSubfolder)
    .addWidget(buttonSet);

  var paymentSection = getPaymentSection();
  var cardHeader = CardService.newCardHeader()
    .setTitle("DriveWorks")
    .setSubtitle("Drive operations made easy")
    .setImageUrl(getLogoURL());
  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(paymentSection)
    .addSection(mainSection);

  return card.build();  
}

