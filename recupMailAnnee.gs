//const id2023 = '1E6aP4Xn1gtg_GaLen4T3yQdkdWY-7C664Hdj-H-Ss1U'
const id2024 = '1kCwBnkcOg6HzS5dT6hReYtczTZH5vwGVnQ_gU-Nk5zE'

const label = "Resa2022";
function recupDataMailIA2024() {
  Logger.log("Lancement V4 Recup Mail");

  const listeRef = ss.getSheetByName("Resa2024").getDataRange().getValues();

  listeRef.forEach((ref,index) => {
    if (ref[1] != "OK") {
      var threads = GmailApp.search(ref[0]);
      var msgs = GmailApp.getMessagesForThreads(threads);
      for (i = msgs.length -1; i > -1; i--) {
        const objResa = callMistralAI (msgs[i][0].getPlainBody());
          objResa.DateModif = msgs[i][0].getDate()
          console.log(objResa);
          if (objResa.Statut != "Erreur") traiterResa(objResa, param,"dataMail3");
      }
      ss.getSheetByName("Resa2024").getRange(index+1,2).setValue("OK");
    }
  });
}

function recupDataMailPeriod() {
  // Initialisation des tableaux de paramètres
  const ssAnnee = SpreadsheetApp.openById(id2024) //  pour remplissage  

  const sheet = ssAnnee.getSheetByName("dataMail");
  const sheetError = ssAnnee.getSheetByName("Erreur");

  Logger.log("Lancement V1 Recup Mail sur une periode");
  var annee=2024;
  for (month = 1; month<2; month++) { 
//    const request = 'label:Save before:' + month + '/1/' + annee + ' after:' + (month-1) + '/1/' + annee
    const request = 'label:Save before:' + 1 + '/1/' + 2025 + ' after:' + 1 + '/12/' + 2024
    Logger.log(request)
    var threads = GmailApp.search(request);

//    var threads = GmailApp.search('label:record -"vous avez un nouveau message"');
    var msgs = GmailApp.getMessagesForThreads(threads);

    var dataExistante = sheet.getDataRange().getValues();

  //  if (sheet.getFilter() != null) sheet.getFilter().remove();

    for (i = msgs.length -1; i > -1; i--) traiterMessage (msgs[i][0], param, sheetError, dataExistante);
    sheet.getDataRange().clearContent();
    var range = sheet.getRange(1,1,dataExistante.length,dataExistante[0].length);  
    range.setValues(dataExistante);

    for (i = threads.length -1; i > -1; i--) {
      var label = GmailApp.getUserLabelByName("Resa2024");
      var label2 = GmailApp.getUserLabelByName("Save");
      label.addToThread(threads[i]);
      label2.removeFromThread(threads[i]);
    }

    Logger.log("TRAITEMENT OK OK pour année " + annee + ", mois " + month);
  }
}

function recupDataMailReplace() {
  // Initialisation des tableaux de paramètres

  const sheet = ss.getSheetByName("dataMail");
  const sheetError = ss.getSheetByName("Erreur");

  var dataExistante = sheet.getDataRange().getValues();

  Logger.log("Lancement recupMail pour Remplacement");
  listeRef.forEach(ref => {
    var threads = GmailApp.search(ref);
    var msgs = GmailApp.getMessagesForThreads(threads);
    for (i = msgs.length -1; i > -1; i--) traiterMessage (msgs[i][0], param, sheetError, dataExistante,true);

  });
  sheet.getDataRange().clearContent();
  var range = sheet.getRange(1,1,dataExistante.length,dataExistante[0].length);  
  range.setValues(dataExistante);
}



function traiterMessage (msgTemp, param, sheetError, dataExistante,flagReplace) {
//-------------------------------------------------------------------------------------------//  
//--------------------------- Début Boucle sur les messages ---------------------------------//
//-------------------------------------------------------------------------------------------//
  const correspondance   = param.correspondance   ;
  const tabTypeCh   = param.tabTypeCh   ;
  const tabTypeRate = param.tabTypeRate ;
  const tabTypePromo= param.tabTypePromo;
  const tabTypeSup  = param.tabTypeSup  ;
  const tabTypeCom  = param.tabTypeCom  ;
  const tabPeriode  = param.tabPeriode  ;
  const tabContrat  = param.tabContrat  ;

  var body = msgTemp.getBody();
  body = body.replace (/\n/g,' '); 
  body = body.replace (/\r/g,'');
  Logger.log("Traitement du message : " + msgTemp.getSubject());

  var body2 = msgTemp.getPlainBody(); // sans les balises html
//    Logger.log(body2);
  
  var valDateRec = msgTemp.getDate();//date resa
  var valEtat = traduireTexte (msgTemp.getSubject(),correspondance); // etat
  var valSource = chercheTexte (body2, "Source :", " ", "*"); //source
  
  var valRefSource = chercheTexte (body2, "#", "", "\n"); //référence du fournisseur BOOKING, HOTEL BEDS ...
  var valRefReseliva = chercheTexte(body2, 'Numéro de', "*", "*"); //ref reseliva 
  if (valSource == "Erreur") { valSource = "Reseliva"; valRefSource = valRefReseliva;} // Pour les resa RESELIVA il n'y a pas la source
  var indexExist = 0;
  var valNbChInitial = 0;
  var valTitulaire = chercheTexte(body2, 'Titulaire de la réservation', ' ', '\n'); // Nom du client
  
  var valDateDeb = strToDate(chercheTexte(body2, 'Date d’arrivée', ' ', '\n'));
  var valDateFin = strToDate(chercheTexte(body2, 'Date de départ', ' ', '\n'));
  var valNbNuit = Math.round((valDateFin-valDateDeb)/86400000);

  if (valTitulaire == "Riad Ambre & Epices ") { // si réservartion enregistré par le Riad, le nom du client est dans le champs informations
    valTitulaire = chercheTexte(body2, 'Informations du client', ' ', '\n');
    valSource += " Pannel"; 
  }
  


  for (k=0 ; k< dataExistante.length ; k++) {
    if (dataExistante[k][1] == valRefReseliva) { 
      if (indexExist == 0) indexExist = k;
      valNbChInitial++;
    }
  }

  if ((valEtat == "Confirmation") && flagReplace) valEtat = "Modification";

  if ((valEtat == "Annulation") && (indexExist != 0)) {
    k = indexExist ;
    if (dataExistante[k][2] != "Annulation") 
      while ((k<dataExistante.length) && (dataExistante[k][1] == valRefReseliva)) { dataExistante[k][2] = valEtat; for (var a=12; a< 20; a++) dataExistante[k][a]=0 ; 
                                                                                    dataExistante[k][20] += ("// Reservation le " + dataExistante[k][0]); 
                                                                                    dataExistante[k][21] = valDateRec.toLocaleDateString("fr-FR"); k++;
                                                                                  }
  }
  else if (((valEtat == "Confirmation") && (indexExist == 0)) || (valEtat == "Modification")) { // Si une confirmation est traité après une annulation on ne fait rien
    try {
      // Création du tableau des chambres tel qu'il est dans le mail 
      var tabChambre = recupTableauChambre(body);

      if (tabChambre != "noTabChambre") { // S'il n'y a pas de tableau de chambre dans le message on ne fait rien
        
        // réupération de la ligne avec le total de la chambre ==> Colone Total Chambre
        var noTotalLigne = 1, nossTotalLigne =1, nol =0;
        while ((nol < tabChambre.length) && (typeof(tabChambre[nol][1]) != "undefined")) {
          if (tabChambre[nol][1].indexOf ('Total pour la') != -1) noTotalLigne = nol; // Identfication du positionnement du total par chambre
          if (tabChambre[nol][1].indexOf ('Sous-total') != -1) nossTotalLigne = nol; // Identfication du positionnement du sous total pour calculer le prix net de vente
          if (tabChambre[nol][1].indexOf ('Total apr') != -1) nossTotalLigne = nol; // Identfication du positionnement du total après remises (cas RESELIVA) pour calculer le prix net de vente
          nol++;
        }
        if (nossTotalLigne == 1)  nossTotalLigne = noTotalLigne; // cas où on ne trouve rien
                
        var valNbChambre = recupNbChambre (body2); // on va récupérer le nombre de chambre car nous créerons 1 ligne par chambre

        var valTotalG = parseFloat(chercheTexte(body2,'*Total général','* ',' EUR'));
        var valInfoC = "";
        if (valSource=="HotelBeds") valInfoC = chercheTexte(body2,'Booking ID:','*','Pour voir ').replace(/\n/g,' '); // interessant uniquement pour les Hotel Beds
        
        var valDateRecText = valDateRec.toLocaleDateString("fr-FR"); // cela permet de supprimer l'heure
        if (valEtat == "Modification") {
          if (indexExist !=0) valDateRecText = dataExistante[indexExist][0]; // en cas de modification, nous mettrons la date de création
          valInfoC += "Modification le " + valDateRec.toLocaleDateString("fr-FR");
        }
        var vecteur = [];
        
        //-------------------------- DEBUT 2EME BOUCLE sur le Nombre de Chambre -------------------------//      
        for (j = 0 ; j< valNbChambre ; j++) {        
          vecteur[j] = [valDateRecText,valRefReseliva,valEtat,valSource,valRefSource, valTitulaire,valDateDeb,valDateFin,valNbNuit,
                        searchV(tabChambre[0][2*j+1].substring(9),tabTypeCh), // Initiale de la chambre réduit
                        searchVPartiel(tabChambre[1][2*j+1],tabTypeRate,2), // Service
                        searchVPartiel(tabChambre[1][2*j+1],tabTypeRate,1), // Contrat
                        nbPaxPayant(tabChambre[2][2*j+1]), // PAX hors enfants gratuits
                        promo (tabChambre[1][2*j+1],valInfoC,tabTypeRate,tabTypePromo),0,0,0,0,0,0, // Promo, le reste (tarifs) est calculé après
                        valInfoC + "/ Contrat :" + tabChambre[1][2*j+1],valDateRecText]; // Autres infos du mail + le contrat pour éventuel analyse d'erreur
          
          var valSup = searchV(searchVPartiel(tabChambre[1][2*j+1],tabTypeRate,2), tabTypeSup)*vecteur[j][12];
          var valTypeContrat = searchVPartiel(tabChambre[1][2*j+1],tabTypeRate,4);
          vecteur[j][16] = valSup;
          if ( valTypeContrat == "B2C") { // Dans ce cas l'information du mail est totalement valide, il faut supprimer le montant des suppléments pour le prix par nuitée
            vecteur[j][14] = parseFloat(tabChambre[nossTotalLigne][2+2*j])/valNbNuit -valSup * (1-vecteur[j][13]); // la promo s'applique au supplément, récupérer le sous-total
            vecteur[j][15] = vecteur[j][14]*(1 - searchV(valSource,tabTypeCom));
            vecteur[j][17] = vecteur[j][15] * valNbNuit; // Calcul du total Nuitée 
            vecteur[j][18] = valSup * valNbNuit; // Calcul du total Supplément 
            vecteur[j][19] = valTotalG;
          }
          if ( valTypeContrat == "BAR") { // Dans ce cas, le tarif du mail n'intègre pas les suppléments
            vecteur[j][14] = parseFloat(tabChambre[nossTotalLigne][2+2*j])/valNbNuit; 
            vecteur[j][15] = vecteur[j][14]*(1 - searchV(valSource,tabTypeCom));
            vecteur[j][17] = vecteur[j][15] * valNbNuit; // Calcul du total Nuitée 
            vecteur[j][18] = valSup * valNbNuit; // Calcul du total Supplément 
            vecteur[j][19] = "NC" ; 
          }
          if ( valTypeContrat == "FIT") { // Dans ce cas, le tarif est calculé
            vecteur[j][15] = tarifFIT (valDateRec,valDateDeb,vecteur[j][9], vecteur[j][11],tabContrat, tabPeriode) * (1-vecteur[j][13]); 
            vecteur[j][17] = (vecteur[j][15] + valSup * (1-vecteur[j][13]))* valNbNuit ;// la promo s'applique au supplément
            vecteur[j][18] = valSup * valNbNuit; // Calcul du total Supplément 
            vecteur[j][19] = "NC"; 
          }
        }
        
        //-------------------------- FIN 2EME BOUCLE sur le Nombre de Chambre -------------------------//      
        
        // On ajoute pour une confirmation et on annule et remplace pour une modification (fonction splice) directement dans le tableau des données récupéré
        if ((valEtat == "Confirmation") || (indexExist ==0)) {
          Logger.log("Ajout de la réservation " + valRefReseliva + " pour nb de chambre "  + valNbChambre);
          for (j = 0 ; j< valNbChambre ; j++) dataExistante.push(vecteur[j]);
        }
        if ((valEtat == "Modification") && (indexExist !=0)) {
          dataExistante.splice(indexExist, valNbChInitial);
          Logger.log("Modification de la réservation " + valRefReseliva);
          for (j = 0 ; j< valNbChambre ; j++) dataExistante.splice(indexExist,0,vecteur[j]);
        }
      }
    }
    catch (error) {
      Logger.log(error + " reférence " + valRefReseliva);
      const now = new Date()
      sheetError.appendRow([now,valRefReseliva, error]);
      return false
    }
  }
  index = 0;
  return true
}

const listeRef= [
6353083,
6430176,
6430176,
6430176,
6430176,
6430202,
6430304,
6430304,
6430304,
6459740,
6526447,
6526447,
6554923,
6561584,
6691593,
6726306,
6733427,
6737974,
6751559,
6752345,
6758754,
6771596,
6775824,
6775808,
6775777,
6778857,
6789431,
6822388,
6829751,
6830528,
6959491,
6959496,
7007138,
7019562,
7063367,
7076137,
7192474,
7224191,
7224191,
7286106]
