function onHomepage(e) {
  console.log('onHomePage = ' + JSON.stringify(e));
  return createHomeCard();
}

function applyDeleteOnFile(file, dryrun) {
  if (dryrun) {
    console.log('Deleting file ' + file.getName());
    return;
  }

  file.setTrashed(true);
}

function applyDeleteOnFolder(folder, dryrun, delete_ops) {
  if (delete_ops.delete_empty_folder) {
    var isEmpty = folder.getFiles().hasNext();
    if (!isEmpty) return;
  }
  if (dryrun) {
    console.log('Deleting folder ' + folder.getName());
    return;
  }

  folder.setTrashed(true);
}

function getNewName(name, rename_ops) {
  if (rename_ops.rename_method === "rename_partial") {
    return name.replaceAll(search, replace);
  } 
  
  if (rename_ops.rename_method === "rename_full") {
    return rename_ops.fullname;
  } 
  
  if (rename_ops.rename_method === "rename_adding") {
    return (rename_ops.before === null)? "" : rename_ops.before + name + (rename_ops.after === null)? "" : rename_ops.after;
  } 
  
  throw "Unsupported rename ops = " + rename_ops.rename_method;  
}

function applyRenameOnFile(file, dryrun, rename_ops) {
  if (dryrun) {
    console.log('Renaming file ' + file.getName());
    return;
  }

  var new_name = getNewName(file.getName(), rename_ops);
  file.setName(new_name);
}

function applyRenameOnFolder(folder, dryrun, rename_ops) {
  if (dryrun) {
    console.log('Renaming folder ' + folder.getName());
    return;
  }

  var new_name = getNewName(folder.getName(), rename_ops);
  folder.setName(new_name);  
}

function nextIteration(iterationState, operation, entity, include_subfolder, dryrun, delete_ops, rename_ops) {
  var currentIteration = iterationState[iterationState.length-1];
  if (currentIteration.fileIteratorContinuationToken !== null) {
    var fileIterator = DriveApp.continueFileIterator(currentIteration.fileIteratorContinuationToken);
    if (fileIterator.hasNext()) {
      var file = fileIterator.next();
      if (entity === "file") {
        if (operation === "delete") {
          applyDeleteOnFile(file, dryrun);
        } else if (operation === "rename") {
          applyRenameOnFile(file, dryrun, rename_ops);
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
          applyDeleteOnFolder(folder, dryrun, delete_ops);
        } else if (operation === "rename") {
          applyRenameOnFolder(folder, dryrun, rename_ops);
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

/*
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

  if (type === "pdf") return "application/pdf";
  if (type === "bmp") return "image/bmp";
  if (type === "gif") return "image/gif";
  if (type === "jpeg") return "image/jpeg";
  if (type === "png") return "image/png";
  if (type === "svg") return "image/svg+xml";
  if (type === "css") return "text/css";
  if (type === "csv") return "text/csv";
  if (type === "html") return "text/html";
  if (type === "js") return "application/javascript";
  if (type === "txt") return "text/plain";
  if (type === "rtf") return "text/rtf";
  if (type === "zip") return "application/zip";

*/
function getMimeType(ops) {
  var type = ops.file_type;
  if (type === "all") return null;
  if (type === "spreadsheet") return "application/vnd.google-apps.spreadsheet";
  if (type === "doc") return "application/vnd.google-apps.document";
  if (type === "slide") return "application/vnd.google-apps.presentation";
  if (type === "form") return "application/vnd.google-apps.form";
  if (type === "sites") return "application/vnd.google-apps.site";
  if (type === "drawing") return "application/vnd.google-apps.drawing";
  if (type === "appscript") return "application/vnd.google-apps.script";
  throw "Unsupported mime type = " + type;

}

function getSearchToken(entity, mimeType, matchStr) {
    var searchStr = null;
    if (mimeType !== null) {
      searchStr = "mimeType = " + mimeType;
    }
    if (matchStr !== null) {
      if (searchStr !== null) {
        searchStr += " and ";
      }
      searchStr += "name contains '" + matchStr + "'";
    }
      
    if (searchStr !== null) {
      console.log('Get token via search string = ' + searchStr);
      if (entity === "file") {
        return folder.searchFiles(searchStr).getContinuationToken();
      }

      return folder.searchFolders(searchStr).getContinuationToken();
    }

    return null;
}

function makeIterationFromFolder(folder, operation, entity, delete_ops, rename_ops) {
  var iteration = {
    folderName: folder.getName(), 
    fileIteratorContinuationToken: folder.getFiles().getContinuationToken(),
    folderIteratorContinuationToken: folder.getFolders().getContinuationToken()
  };

  if (operation === "delete") {
    if (entity === "file") {
      console.log('mime type = ' + getMimeType(delete_ops) + ', match str = ' + getMatchStr(delete_ops));
      var token = getSearchToken(entity, getMimeType(delete_ops), getMatchStr(delete_ops));
      if (token !== null) {
        console.log('search file token is not null = ' + token);
        iteration.fileIteratorContinuationToken = token;
      }
    } else {
      var token = getSearchToken(entity, null, getMatchStr(delete_ops));
      if (token !== null) {
        iteration.folderIteratorContinuationToken = token;
      }
    }
  } else if (operation === "rename") {
    if (entity === "file") {
      var token = getSearchToken(entity, getMimeType(rename_ops), getMatchStr(rename_ops));
      if (token !== null) {
        iteration.fileIteratorContinuationToken = token;
      }
    } else {
      var token = getSearchToken(entity, null, getMatchStr(rename_ops));
      if (token !== null) {
        iteration.folderIteratorContinuationToken = token;
      }
    }
  }

  return iteration;
}

function iterateFolder(folder, operation, entity, include_subfolder, dryrun, delete_ops, rename_ops) {
  var email = Session.getEffectiveUser().getEmail();
  console.log('Iterate entity in folder: email = ' + email + ', folder = ' + folder + ', operation = ' + operation + ', entity = ' + entity + ', include_subfolder = ' + include_subfolder + ', dryrun = ' + dryrun + ', delete_ops = ' + JSON.stringify(delete_ops) + ', rename_ops = ' + JSON.stringify(rename_ops));
  var MAX_RUNNING_TIME_MS = 4.5 * 60 * 1000;
  var startTime = (new Date()).getTime();
  var iterationState = JSON.parse(PropertiesService.getUserProperties().getProperty(email));
  if (iterationState !== null) {
    if (folder.getName() !== iterationState[0].folderName) {
      console.error("Iterating a new folder: " + folder.getName() + ". End early since existing operation is not done.");
      return;
    }
    console.info("Resuming iteration for folder: " + folder.getName());
  }
  if (iterationState === null) {
    console.info("Starting new iteration for folder: " + folder.getName());
    iterationState = [];
    iterationState.push(makeIterationFromFolder(folder, operation, entity, delete_ops, rename_ops));
  }  

  while (iterationState.length > 0) {
    iterationState = nextIteration(iterationState, operation, entity, include_subfolder, dryrun, delete_ops, rename_ops);
    var currTime = (new Date()).getTime();
    var elapsedTimeInMS = currTime - startTime;
    var timeLimitExceeded = elapsedTimeInMS >= MAX_RUNNING_TIME_MS;
    if (timeLimitExceeded) {
      PropertiesService.getUserProperties().setProperty(email, JSON.stringify(iterationState));
      console.info("Stopping loop after '%d' milliseconds.", elapsedTimeInMS);
      return;
    }
  }

  console.info("Done iterating. Deleting iterating state ... ");
  PropertiesService.getUserProperties().deleteProperty(email);
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

function deleteFileHandler(e) {
  console.log('deleteFileHandler = ' + JSON.stringify(e));
  var folder = parseFolderFromEvent(e);
  var filename_match = ('file_name_field' in e.formInput)? e.formInput.file_name_field : null;
  var delete_ops = {
    file_type: e.formInput.file_type_field,
    search: filename_match,
    delete_empty_folder: null,
  }
  iterateFolder(folder, "delete", "file", JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun), delete_ops, null);
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
  iterateFolder(folder, "delete", "folder", JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun), delete_ops, null);
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
  iterateFolder(folder, "rename", "file", JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun), null, rename_ops);
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
  iterateFolder(folder, "rename", "folder", JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun), null, rename_ops);
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
  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(paymentSection)
    .addSection(mainSection);

  return card.build();  
}

function createDeleteFolderCard(include_subfolder, dryrun) {
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
  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(paymentSection)
    .addSection(mainSection);

  return card.build();  
}

function configureMore(e) {
  console.log('configureMore = ' + JSON.stringify(e));
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

function getPaymentSection() {
  var endIcon = CardService.newIconImage().setIcon(CardService.Icon.VIDEO_PLAY);
  var buymoreIcon = CardService.newIconImage().setIcon(CardService.Icon.DOLLAR);
  var buymore = CardService.newDecoratedText().setText("Remaining 7 days for trial").setStartIcon(buymoreIcon).setEndIcon(endIcon).setWrapText(true);
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
    .addItem("Preview operations only without executing", "dryrun", true);

  var includeSubfolder = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("include_subfolders_field")
    .addItem("Apply to subfolders", "include_subfolder", false);

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

