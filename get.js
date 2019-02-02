function getEtudeInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  etude.nom = getCellValue(sheet, 0, 'B3');
  etude.nomFormate = formatNomEtude(etude.nom);

  etude.dateDebut = getCellValue(sheet, 0, 'B21');
  etude.semaineFin = getCellValue(sheet, 0, 'B23');

  etude.dateRef = Utilities.formatDate(etude.dateDebut, "GMT+2", "yyMMdd")

  etude.dateDebut = Utilities.formatDate(etude.dateDebut, 'GMT+2', 'dd/MM/yyyy');

  etude.clientSociete = getCellValue(sheet, 0, 'F11');

  etude.sujet = getCellValue(sheet, 0, 'B5');

  etude.ce = getCellValue(sheet, 0, 'B27');
}

function getAdminInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  responsableProjet.sexe = getCellValue(sheet, 0,'E10');
  responsableProjet.nom = getCellValue(sheet, 0,'D10');
  responsableProjet.prenom = getCellValue(sheet, 0,'B10');
  responsableProjet.portable = getCellValue(sheet, 0,'F10');
  responsableProjet.email = getCellValue(sheet, 0,'G10');
}

function getPresidentInfo(sheet) {
  president.sexe = getCellValue(sheet, 0, 'E3');
  president.nom = getCellValue(sheet, 0, 'B3');
  president.prenom = getCellValue(sheet, 0, 'A3');
  president.portable = getCellValue(sheet, 0, 'C3');
  president.email = getCellValue(sheet, 0, 'D3');
}

function getEtudiant(sheet) {
  var sheetActive = SpreadsheetApp.getActiveSheet();

  var range = sheetActive.getRange("B2:B10");
  var IDEtudiant = range.getCell(1,1).getValue();

  var rangePoste = sheetActive.getRange("B3:B10");
  etudiant.poste = rangePoste.getCell(1,1).getValue();

  var rangeDescription = sheetActive.getRange("B4:B10");
  etudiant.descriptionMissions = rangeDescription.getCell(1,1).getValue();

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] == IDEtudiant) {
        var rowID = i+1;

        etudiant.nom = getCellValue(sheet, 0, "F"+rowID);
        etudiant.prenom = getCellValue(sheet, 0, "G"+rowID);
        etudiant.sexe = getCellValue(sheet, 0, "E"+rowID);
        etudiant.adresse = getCellValue(sheet, 0, "V"+rowID);
        etudiant.codePostal = getCellValue(sheet, 0, "X"+rowID);
        etudiant.ville = getCellValue(sheet, 0, "W"+rowID);
        etudiant.portable = getCellValue(sheet, 0, "O"+rowID);
        etudiant.mail = getCellValue(sheet, 0, "H"+rowID);
      }
    }    
  }

  if(etudiant.sexe  == 'M') {
    etudiant.pronom = 'Il'; 
  } else if (etudiant.sexe == 'F') {
    etudiant.pronom = 'Elle';
  }
} 

function getInfosEtudiant() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var rangeMontantJEH = sheet.getRange("B8:B10");
  var rangeNbJEH = sheet.getRange("B7:B9");
  var rangeRemuneration = sheet.getRange("B9:B11");

  etudiant.montantJEH = rangeMontantJEH.getCell(1,1).getValue();
  etudiant.nbJEH = rangeNbJEH.getCell(1, 1).getValue();
  etudiant.remuneration = rangeRemuneration.getCell(1,1).getValue();

}