const idRecupMail = "1jxvJt9GL3dSUsANBr2cFhdCv8tbe2_vNTOXckxxN0Bo";
const idManageApp = "1hv3F_DZmIJB1eILI5vLJrEQ_iPDbXXm3UZ_taH0yJ9E";
// const idFicheApp = "1U6XAjRsnwKDFSUBri_KHipzyeGSIn-NSSi4ynolb3TU"; // Version officielle
const idFicheApp = "12Xr19UvZ1sUDNrR1kJVUFgIBEEHwBdSGgILLWZvUQgw"; // Version officielle

var room =["ST","DE","SU","SX","SF"];

function getClientFutur() {
  const now = new Date();
  var tabReturn = SpreadsheetApp.openById(idRecupMail).getSheetByName("dataMail").getDataRange().getValues().filter(
    tab => (tab[2] == "Confirmation") && (tab[6] > now));
  tabReturn = tabReturn.sort((a, b) => a[6] - b[6]);
  
  var tabDataUnique = tabReturn.filter((r, i, t) => (t.findIndex(r2 => r2[1] == r[1]) === i));  
  var tabResa = []; // init tableau d'objet
  for (var i = 0; i< tabDataUnique.length; i++) { // en cas de resa multiples chambres on va ajouter les paxs et concatener les types de chambres
    tabDataUnique[i][0] = tabDataUnique[i][0].toLocaleDateString("fr-FR");
    tabDataUnique[i][6] = tabDataUnique[i][6].toLocaleDateString("fr-FR");
    tabDataUnique[i][7] = tabDataUnique[i][7].toLocaleDateString("fr-FR");
    tabResa[i] = createResaFromRecupMail(tabDataUnique[i]);
    var tabDataIdem = tabReturn.filter (r => r[1] == tabResa[i].idResa);

    tabResa[i].nbRoom = tabDataIdem.length;
    tabDataIdem.forEach(r => tabResa[i].Pax += r[12]);
    for (j = 0; j<room.length; j++) {
      var nb = tabDataIdem.filter(r => r[9] == room[j]).length;
      if (nb != 0) tabResa[i].Room += (" " +nb + " " + room[j]);
    }
  }
  return tabResa;
}

function createFicheClient (resa) {
  const ss = SpreadsheetApp.getActive();
  var invoicePage = ss.getSheetByName("Invoice");
  invoicePage.appendRow(resaToTab(resa));
}

//////////////////////////////////////
////////////// Récupération du planning
//////////////////////////////////////
function getPlanning () {
  const ssPlanning = SpreadsheetApp.openById(idManageApp);
  var mMoins1 = new Date();
  mMoins1.setDate(mMoins1.getDate() - 60);

  var tabPlanning = ssPlanning.getSheetByName("Planning").getDataRange().getValues().filter(r => r[2] > mMoins1);

  tabPlanning.shift();
  tabPlanning.forEach(function(r) {
    r[1]=r[1].getTime();
    r[2]=r[2].getTime();
    if (r[3] != "") r[3]=r[3].getTime();
  });
  var lastUpdate = ssPlanning.getSheetByName('Param').getRange('O2').getValue(); // Recup de la dernière update sauvegardé
  ssPlanning.getSheetByName('Param').getRange('O3').setValue(lastUpdate); // Mise à jour de la date du dernier update sauvegardé en cas de cancel
  return tabPlanning;
}

//////////////////////////////////////
////////////// Sauvegarde du planning
//////////////////////////////////////
function setPlanning (tabPlanning) {
  const ssPlanning = SpreadsheetApp.openById(idManageApp);
  const sheetPlanning = ssPlanning.getSheetByName("Planning");
  
  tabPlanning.forEach(function(r) { // fusion des même resas sur même ligne séparé
    var index = tabPlanning.findIndex (r2 => (r2[9] == r[9]) && (r2[8] == r[8]) && (r2[1] == r[2]));
    if (index != -1) {    
      r[2] = tabPlanning[index][2];
      tabPlanning[index][8] = "Delete";
    }
  });
  
  tabPlanning = tabPlanning.filter(r => r[8] != "Delete").sort((a, b) => a[1] - b[1]); // deletion of resa Delete by the user or by the system in loadNewResa
  var tabReturn = [];
  tabPlanning.forEach(function(r) {// change time in date, construct the tab to return and Delete the statute information
    if (r[8] == "to affect") r[10] = "to affect";
    else r[10] = "";
    tabReturn.push(r.filter(r2 => true));
    r[1]= new Date (r[1]);
    r[2]= new Date (r[2]);
    if (r[3] != "") r[3] = new Date (r[3]);
  });
  sheetPlanning.getRange(2,1, sheetPlanning.getLastRow(), sheetPlanning.getLastColumn()).clear();
  var range = sheetPlanning.getRange(2,1, tabPlanning.length, tabPlanning[0].length);
  range.setValues(tabPlanning);
  var lastUpdate = ssPlanning.getSheetByName('Param').getRange('O3').getValue(); // Recup de la dernière update avant sauvegarde
  ssPlanning.getSheetByName('Param').getRange('O2').setValue(lastUpdate); // Mise à jour de la date du dernier update vu qu'il y a sauvegarde
  return tabReturn;
 }

//////////////////////////////////////
////////////// Récupération des dernieres reservation non intégré dans le planning
//////////////////////////////////////
function loadNewResa () { // mise à jour du planning en récupérant les resa non ajoutées et supprimant les annulations
  const now = new Date();
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 2);
  
  const ss = SpreadsheetApp.openById(idRecupMail);
  const ssPlanning = SpreadsheetApp.openById(idManageApp);
  const sheetPlanning = ssPlanning.getSheetByName("Planning")
  const lastUpdate = ssPlanning.getSheetByName('Param').getRange('O2').getValue();
  var filtreLastUpdate = addDays(lastUpdate, -1);
  var tabPlanning = sheetPlanning.getDataRange().getValues();
  tabPlanning.shift();
  var listeRef = tabPlanning.map(r => r[9]); 
  
// récupération des mails et du tableau constitué B2C
  // var tabRecupMail = TraitementAuto.recupDataMail().filter(tab => (tab[3] != "airbnb") && (tab[7] > now) && ((tab[0]>=lastUpdate) || (tab[21]>=lastUpdate))).reverse();
  var tabRecupMail = ss.getSheetByName("dataMail").getDataRange().getValues().filter(r => r[7] > yesterday);
  
// on cherche les resa dans recup qui ne sont pas dans le planning, on s'appuie sur les références de resa, si oui on les ajoute au planning avec "to affect" comme chambre
// récupération des mails et du tableau constitué B2C
  var refTraite = ""; // pour éviter de traiter plusieurs fois une mếmé référence
  tabRecupMail.forEach(function (r) {
    if (refTraite != r[1]) {
      const tabPlId = tabPlanning.filter(r2 => r2[9] == r[1]); // recup de tous les reservation avec le même Id
      const tabRecupId = tabRecupMail.filter(r2 => r2[1] == r[1]); // nb Chambre reservé
      var testModifImportante = ((r[2] == "Modification") && (tabPlId.length == 0)) || ((r[2] == "Modification") // une des dates change, le nombre de chambre change
        && ((tabPlId[0][1].getTime() != r[6].getTime()) || (tabPlId[0][2].getTime() != r[7].getTime()) || (tabRecupId.length != tabPlId.length)))

      if ((r[2] == "Annulation") || testModifImportante)  { // Pour les grosses modifs on supprime et on remet
        tabPlanning.forEach(r2 => r2[9] == r[1] ? (r2[8] = "Delete") &&  (r2[10] = "Delete") : null);
      }

      if (((r[2] == "Confirmation") && (tabPlId.length == 0)) || testModifImportante) // confirmation pas dans le planning ou modif nécessitant réédition
        tabRecupId.forEach(r2 => tabPlanning.push([r2[5], r2[6], r2[7], r2[0], r2[3],"",r2[19], r2[9],"to affect", r2[1], testModifImportante ? "Modif" : "New"]));
         // ajout des resas confirmés
      refTraite = r[1]
    }
  });

  ssPlanning.getSheetByName('Param').getRange('O3').setValue(now); // Mise à jour de la date du dernier update
  tabPlanning.forEach(function(r) {
    r[1]=r[1].getTime();
    r[2]=r[2].getTime();
    if (r[3] != "") r[3]=r[3].getTime();
  });
  return tabPlanning;
}

//////////////////////////////////////
////////////// Récupération des param
//////////////////////////////////////
function getUpdate() {
  const ssPlanning = SpreadsheetApp.openById(idManageApp);
  const lastUpdate = ssPlanning.getSheetByName('Param').getRange('O2').getValue();
  return lastUpdate.getTime();
}

function addDays(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

//// Check in client
function checkIn (guestInfo) { // Mise à jour de l'autre fichier de suivi des clients
  return Manageappcode.checkIn(guestInfo);
}
