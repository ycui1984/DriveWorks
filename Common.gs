function onHomepage(e) {
  console.log('onHomePage = ' + JSON.stringify(e));
  var properties = PropertiesService.getUserProperties();
  var email = Session.getEffectiveUser().getEmail();
  properties.deleteProperty(email);
  return createHomeCard();
}

function iterateFolder(folder, operation, entity, include_subfolder, dryrun, target_file_type) {
  console.log('iterate entity in folder: folder = ' + folder + ', operation = ' + operation + ', entity = ' + entity + ', include_subfolder = ' + include_subfolder + ', dryrun = ' + dryrun + ', target_file_type = ' + target_file_type);
  
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
  iterateFolder(folder, "delete", "file", JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun), e.formInput.file_type_field);
}

function deleteFolderHandler(e) {
  console.log('deleteFolderHandler = ' + JSON.stringify(e));
  var folder = parseFolderFromEvent(e);
  iterateFolder(folder, "delete", "folder", JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun), null);
}

function renameFileHandler(e) {
  console.log('renameFileHandler = ' + JSON.stringify(e));  
  var folder = parseFolderFromEvent(e);
  iterateFolder(folder, "rename", "file", JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun), e.formInput.file_type_field);
}

function renameFolderHandler(e) {
  console.log('renameFolderHandler = ' + JSON.stringify(e));
  var folder = parseFolderFromEvent(e);
  iterateFolder(folder, "rename", "folder", JSON.parse(e.parameters.include_subfolder), JSON.parse(e.parameters.dryrun), null);
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
    .addItem("Delete folders with no files only", "delete_empty_folder", true);
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

