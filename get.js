function getEtudeInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  etude.nom = getCellValue(sheet, 0, 'B3');
  etude.nomFormate = formatNomEtude(etude.nom);

  etude.dateDebut = getCellValue(sheet, 0, 'B21');
  etude.nbSemaine = getCellValue(sheet, 0, 'C22');
  etude.nbJEH = getCellValue(sheet, 0, 'C28')
  etude.semaineFin = getCellValue(sheet, 0, 'B23');

  etude.dateRef = Utilities.formatDate(etude.dateDebut, "GMT+2", "yyMMdd");

  etude.dateDebut = Utilities.formatDate(etude.dateDebut, 'GMT+2', 'dd/MM/yyyy');

  etude.clientSociete = getCellValue(sheet, 0, 'F11');

  etude.cc = getCellValue(sheet, 0, 'B4');

}

function getPriceInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  prix.totalHT = getCellValue(sheet, 5, 'B5');
  prix.totalVA = getCellValue(sheet, 5, 'B6');
  prix.totalTTC = getCellValue(sheet, 5, 'B7');
  prix.montantHT = getCellValue(sheet, 5, 'B3');
  prix.fraisDossier = getCellValue(sheet, 0, 'B30');
  prix.acompteHT = getCellValue(sheet, 5, 'G4');
  prix.acomptePct = getCellValue(sheet, 5, 'F4');
  prix.acompteTTC = getCellValue(sheet, 5, 'H4');
  prix.resteHT = getCellValue(sheet, 5, 'G16');
  prix.resteTTC = getCellValue(sheet, 5, 'H16');
}

function getClientInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  client.sexe = getCellValue(sheet, 0, 'N11');
  client.prenom = getCellValue(sheet,0, 'D11');
  client.nom = getCellValue(sheet,0, 'E11');
  client.portable = getCellValue(sheet,0, 'L11');
  client.email = getCellValue(sheet,0, 'M11');
  // Did not find the adress, points to empty fields in the client infos on tab 0
  client.adresse = getCellValue(sheet, 0, 'G11');
  client.codePostal = getCellValue(sheet, 0, 'H11');
  client.ville = getCellValue(sheet, 0, 'I11');

}

function getAdminInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  responsableCommercial.sexe = getCellValue(sheet, 0,'F9');
  responsableCommercial.nom = getCellValue(sheet, 0,'D9');
  responsableCommercial.prenom = getCellValue(sheet, 0,'E9');
  responsableCommercial.portable = getCellValue(sheet, 0,'G9');
  responsableCommercial.email = getCellValue(sheet, 0,'H9');
}

function getPresidentInfo(sheet) {
  president.sexe = getCellValue(sheet, 0, 'F8');
  president.nom = getCellValue(sheet, 0, 'D8');
  president.prenom = getCellValue(sheet, 0, 'E8');
  president.portable = getCellValue(sheet, 0, 'G8');
  president.email = getCellValue(sheet, 0, 'D3');
}

function getCreationDate() {
  var creation = new Date();
  var dd = creation.getDate();
  var mm = creation.getMonth() + 1; 
  var yyyy = creation.getFullYear();
  
  if (dd < 10) {
    dd = '0' + dd;
  }
  
  if (mm < 10) {
    mm = '0' + mm;
  }
  
  creation = dd + '/' + mm + '/' + yyyy;

  return creation;
}


