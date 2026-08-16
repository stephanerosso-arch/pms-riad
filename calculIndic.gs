var colFiltre =2;
var donneeFiltre = "Annulation";
var signeFiltre = "notEgal";

function calculStock() {
  var ss = SpreadsheetApp.openById(fichier2018);
  var dataMailRecup = ss.getSheetByName("dataMail").getDataRange().getValues();
  
  var dataTest = dataMailRecup.filter(filtrer);
  
  var test =0;
}

function filtrer(a) {
  if (signeFiltre == "notEgal") if (a[colFiltre]!=donneeFiltre) return true;
  if (signeFiltre == "Egal") if (a[colFiltre]==donneeFiltre) return true;
  if (signeFiltre == "Sup") if (a[colFiltre].getTime()>donneeFiltre.getTime()) return true;
  if (signeFiltre == "Inf") if (a[colFiltre].getTime()<donneeFiltre.getTime()) return true;

  return false;
}