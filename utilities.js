function formatNomEtude(str) {    
    str = str.replace(/\s+/g, '-').toLowerCase();
    return str;
}

function getCellValue(ss, nbTab, range) {
    var sheet = ss.getSheets()[nbTab];
    var cell = sheet.getRange(range);
    
    return cell.getValue();
}