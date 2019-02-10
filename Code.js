function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Générer BC', 'createBC');
  menu.addToUi();
}

var bc = {};
var etude = {};
var responsableCommercial = {};
var president = {};
var prix = {}; 
var client = {}; 

function createBC() {
  var SHEET_ADMIN = SpreadsheetApp.openById('1tjSkg05xJw5GAWw6VIuyAfILuaoE_7JSTwNqDDpvLts');
  var SHEET_ADHERENT = SpreadsheetApp.openById('1VL5hMwccmmtyWX5gaMtRrm-U232-gXPvBa2Wz4z_3jE');
  //var SHEET_ETUDES = SpreadsheetApp.openById('1NU-0iyg7wjBsHNsmypZPOASN1xFK2kM_I3QCLDnX-1M').getSheets()[0];
  //var SHEET_BV = SpreadsheetApp.openById('1lGT0FlFkKRi17nPJQkNWSqQfHCXYkT4mUexx6BEbsgY').getSheets()[0];
  //var SHEET_BC = SpreadsheetApp.openById('1DhKbZtB17PFW8oFCI1Mg8Z0d3RO6gbgnHM3fHwNAnd4').getSheets()[0];
  var DOC_MODEL = DocumentApp.openById('19WpV31e6Jms4z2o7pYuglKsbDlqHh9ZzcFus-UOAI9k');
 
  getEtudeInfo();
  getAdminInfo();
  getPresidentInfo(SHEET_ADMIN);
  getClientInfo();
  getPriceInfo();
  getCreationDate();
    
  bc.name = 'BC_' + etude.dateRef + '_' + etudiant.nom + '_' + etude.nomFormate;

  bc.doc = copyBC(DOC_MODEL.getId(), bc.name);

  fillBc(bc.doc);

  openNewDoc(bc.doc.getId());
}

function copyBC(id, name) {
  var folderID = getParentFolder();
  var destFolder = DriveApp.getFolderById(getBCFolder(folderID));
  var file = DriveApp.getFileById(id).makeCopy(name, destFolder);

  return DocumentApp.openById(file.getId());
}

function getParentFolder(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var folders = file.getParents();
  while (folders.hasNext()){
    return folders.next().getId();
  }
}

function getBCFolder(id) {
  // Log the name of every folder that are children of the current folder and you own and is starred.
var folders = DriveApp.getFolderById(id).searchFolders("title contains '3'");
  while (folders.hasNext()) {
    var folder = folders.next();
    return folder.getId();
  }
}

function openNewDoc(id) {
  var url = 'https://docs.google.com/document/d/' + id;
  var html =
    "<script>window.open('" + url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Open BC');

  fillURL(url);
}