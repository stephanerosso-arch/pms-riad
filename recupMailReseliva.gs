var index = 0;
var body;
//var fichierRecup = "1M8HP9DgIt5DRRtoRfkaTIgo4VKVPdSEG-8zQrDsMS4c";
var fichierRecup = "1jxvJt9GL3dSUsANBr2cFhdCv8tbe2_vNTOXckxxN0Bo";
var fichierHB = "1R8Pi-PPpzscI6r70zLTMbNSyTrDefHuPH2PvmtREBQ4";

var fichier2018  = "1ow9NivKrevDh5jeWZk0xi04J1dKeKk1T8mxflXGsrCo";
var mailEnvoye =0;

const ss = SpreadsheetApp.openById(fichierRecup);
const param = {
    correspondance : ss.getSheetByName("Correspondance").getRange("T2:U30").getValues(),
    tabTypeCh : ss.getSheetByName("Correspondance").getRange("K3:L100").getValues(),
    tabTypeRate : ss.getSheetByName("Correspondance").getRange("F3:J500").getValues(),  
    tabTypePromo : ss.getSheetByName("Correspondance").getRange("V2:X30").getValues(),  
    tabTypeSup : ss.getSheetByName("Correspondance").getRange("Q2:S30").getValues(), 
    tabTypeCom : ss.getSheetByName("Correspondance").getRange("N3:O30").getValues(),  
    tabPeriode : ss.getSheetByName("Correspondance").getRange("A3:D100").getValues(),  
    tabContrat : ss.getSheetByName("Contrat").getDataRange().getValues(),
}

function recupDataMail() {
  // Initialisation des tableaux de paramètres
  const sheet = ss.getSheetByName("dataMail");
  const sheetError = ss.getSheetByName("Erreur");

  Logger.log("Lancement V3 Recup Mail");

  var threads = GmailApp.search('label:toRecord -"vous avez un nouveau message"');
  //  var threads = GmailApp.search('label:toTest -"vous avez un nouveau message"');
  //  threads = threads.slice(-10);
  var msgs = GmailApp.getMessagesForThreads(threads);

  var dataExistante = sheet.getDataRange().getValues();

  if (sheet.getFilter() != null) sheet.getFilter().remove();

  for (i = msgs.length -1; i > -1; i--) traiterMessage (msgs[i][0], param, sheetError, dataExistante);
  sheet.getDataRange().clearContent();
  var range = sheet.getRange(1,1,dataExistante.length,dataExistante[0].length);  
  range.setValues(dataExistante);

  for (i = threads.length -1; i > -1; i--) {
    var label = GmailApp.getUserLabelByName("record");
    var label2 = GmailApp.getUserLabelByName("toRecord");
    label.addToThread(threads[i]);
    label2.removeFromThread(threads[i]);
  }
}



function test() {
  var threadsTemp = GmailApp.search('from:(info@reseliva.com) HALF BOARD NRF BAR '); //  test
  recupDataMail(threadsTemp);
}

function updateData () {
  
}

function chercheTexte(bodyText, chaineDonnee, chaineDeb, chaineFin) { // Cherche à partir d'une position donnée et calcul l'index où l'info est trouvée
  var indexDonnee = bodyText.indexOf(chaineDonnee,index)
  var positionDonnee = indexDonnee+chaineDonnee.length; 
  var positionDeb = bodyText.indexOf(chaineDeb,positionDonnee) + chaineDeb.length; 
  var positionFin = bodyText.indexOf(chaineFin,positionDeb+1); 
  index = positionFin;
  if (indexDonnee < 0) return "Erreur";
  return bodyText.substring(positionDeb,positionFin);
}

function ligneSuivante (bodyText) {
  var positionDeb = bodyText.indexOf('\n',index); 
  var positionFin = bodyText.indexOf('\n',positionDeb+1); 
  index = positionFin;
  return (bodyText.substring(positionDeb+1,positionFin));
}

function traduireTexte (texte,tab,col){
  col = col || 1;
  for (var i = 0; i < tab.length; i++) {
    if (texte.indexOf(tab[i][0]) != -1) return tab[i][col];
  }
  return "Erreur Param";
}

function searchV (texte,tab,col){
  col = col || 1;
  if (texte == "Family Suite  2 Bedrooms") var test = 1 ;
      

  for (var i = 0; i < tab.length; i++) if (texte == tab[i][0]) return tab[i][col];
  
  if (mailEnvoye ==0) {
    GmailApp.sendEmail('guilmour101@gmail.com', 'Erreur Paramètre dans Traitement Auto', texte);
    mailEnvoye++;
  }

  return ("Erreur Param :" + texte);
}

function searchVPartiel (texte,tab,col){
  col = col || 1;
  for (var i = 0; i < tab.length; i++) if (texte.indexOf(tab[i][0]) !=-1) return tab[i][col];
  
  if (mailEnvoye ==0) {
    GmailApp.sendEmail('guilmour101@gmail.com', 'Erreur Paramètre dans Traitement Auto', texte);
    mailEnvoye++;
  }

  return ("Erreur Param :" + texte);
}


function recupNbChambre (texte) {
  var nbCh = 1;
  while (texte.indexOf("Chambre " + (nbCh+1)) != -1) nbCh ++;
  return nbCh;
}

function tarifFIT (dateRec,dateDeb,chambre, contrat, tabCont, tabPer) {
  var per = 3 ;
  var grille = 1 ;
  for (var i = 0; i < tabPer.length;i++) {
    if ((tabPer[i][2] <= dateRec) && (tabPer[i][2] != "")) grille = tabPer[i][3];
    if ((tabPer[i][0] <= dateDeb) && (tabPer[i][0] != "")) per = tabPer[i][1];
  }
  for (var i = 0; i < tabCont.length;i++) if ((tabCont[i][0] == grille) &&  (tabCont[i][1] == contrat) &&  (tabCont[i][4] == chambre)) return tabCont[i][4+per];
  return "Tarif non Trouvé";
}


function nbPaxPayant (texte) {
  var indexAdult = texte.indexOf(" Adulte")
  
  if (indexAdult != -1) var nbAdulte = texte.substring(0,indexAdult);
  else return 2;
  
  var indexKid = texte.indexOf(" Enfant");
  var nbKid = 0 ;
  if (indexKid != -1) nbKid = texte.substring(indexKid-1,indexKid);
  if (nbKid == "t") {
    nbKid = 0;
    indexKid = texte.indexOf(" Enfant",indexKid + 1);
    if (indexKid != -1) nbKid = texte.substring(indexKid-1,indexKid);
  }
  return (parseInt(nbAdulte) + parseInt(nbKid));
}

function promo (contrat,infoClient,tabC, tabP) {
  var promoClient = 1; // 100 % à payer pour le client au début
  for (var i = 0; i < tabP.length; i++) if (infoClient.indexOf(tabP[i][0]) != -1) promoClient = promoClient * (1-tabP[i][2]);
  promoClient = promoClient * (1-searchVPartiel(contrat,tabC,3));
  return (1-promoClient);  
}

function strToDate (texte){
  texte = texte.replace("Janvier", "January");                   	
  texte = texte.replace("Février","February");                
  texte = texte.replace("Mars","March");                 
  texte = texte.replace("Avril","April");
  texte = texte.replace("Mai","May");
  texte = texte.replace("Juin","June");
  texte = texte.replace("Juillet","July");
  texte = texte.replace("Août","August");       
  texte = texte.replace("Septembre","September");
  texte = texte.replace("Octobre","October");
  texte = texte.replace("Novembre","November");
  texte = texte.replace("Décembre","December");
  var dateReturn = new Date (texte);
  return dateReturn;
}

function recupTableauChambre (bodyText) {
  const indTabChambre = bodyText.indexOf('Chambres');
  if (indTabChambre == -1) return "noTabChambre";
  bodyText = bodyText.substring(bodyText.indexOf('Chambres'));
  bodyText = bodyText.substring(bodyText.indexOf('<table')); // calage sur le début de tableau des chambres
  var indexBalise = bodyText.indexOf("<");
  var balise = bodyText.substring(indexBalise,bodyText.indexOf(" "));
  var tabCh = new Array();
  for (remp = 0 ; remp <40; remp++) tabCh[remp] = new Array(20);
//  tabCh[0] = new Array();
  var compteurLigne =-1;
  var compteurCol = 0;
  while (balise != "</table") {
    indexBalise = bodyText.indexOf("<",indexBalise);
    balise = bodyText.substring(indexBalise,Math.min(bodyText.indexOf(" ",indexBalise),bodyText.indexOf(">",indexBalise)));
    if (balise == "<tr") {(compteurLigne++); (compteurCol = 0);};
    if (balise == "<td") {
      var indexFin = bodyText.indexOf("</td>",indexBalise);
      var cellContent = bodyText.substring(indexBalise,indexFin);
      indexBalise = indexFin;
      
      var reg=new RegExp("<.[^>]*>", "gi" );
      var cellText=cellContent.replace(reg, "" );
      cellContent = cellContent.replace(">"," >"); // Dans le cas où il n'y a pas d'espace

      var valNbRow =1;
      var valNbCol =1;

      var testColspan = cellContent.indexOf("colspan");
      var testRawspan = cellContent.indexOf("rowspan");
      if (testRawspan !=-1)  var valNbRow = parseInt(cellContent.substring(testRawspan+8,cellContent.indexOf(" ",testRawspan)));
      if (testColspan !=-1)  var valNbCol = parseInt(cellContent.substring(testColspan+8,cellContent.indexOf(" ",testColspan)));
      
      for (var t =0 ; t<valNbRow; t++) for(var u=0 ; u < valNbCol; u++) {
        if (tabCh[compteurLigne+t][compteurCol+u] === undefined) tabCh[compteurLigne+t][compteurCol+u] = cellText;
        else valNbCol++;
      }
          
      compteurCol +=valNbCol;
    }
    indexBalise++;
  }
  return tabCh;
}