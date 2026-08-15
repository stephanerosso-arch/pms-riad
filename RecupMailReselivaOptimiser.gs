function recupDataMailOptim() {
  const sheet = ss.getSheetByName("dataMail2");
  const sheetError = ss.getSheetByName("Erreur");

  Logger.log("Lancement Optimisé Recup Mail");

  var threads = GmailApp.search('6650165');
  var msgs = GmailApp.getMessagesForThreads(threads);

  var { header, existingReservations } = getExistingReservations(sheet);
  var newData = [];
  var rowsToDelete = [];

  for (var i = msgs.length - 1; i >= 0; i--) {
    var msg = msgs[i][0];
    var refReseliva = chercheTexte(msg.getPlainBody(), 'Numéro de', "*", "*"); 
    if (!refReseliva) continue;

    var newRooms = traiterMessageOpt(msg, param, sheetError, existingReservations);

    if (!newRooms || newRooms.length === 0) continue;

    if (existingReservations[refReseliva]) {
      var oldRooms = existingReservations[refReseliva];

      var oldRoomDetails = oldRooms.map(r => r.row[7]); // Type de chambre
      var newRoomDetails = newRooms.map(r => r[7]); 

      oldRooms.forEach(oldRoom => {
        if (!newRoomDetails.includes(oldRoom.row[7])) {
          rowsToDelete.push(oldRoom.index);
        }
      });

      newRooms.forEach(newRoom => {
        var existingRoom = oldRooms.find(r => r.row[7] === newRoom[7]);

        if (existingRoom) {
          updateExistingRow(sheet, existingRoom.index, newRoom);
        } else {
          newData.push(newRoom);
        }
      });
    } else {
      newData.push(...newRooms);
    }
  }

  if (newData.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newData.length, newData[0].length).setValues(newData);
  }

  if (rowsToDelete.length > 0) {
    deleteRows(sheet, rowsToDelete);
  }

  Logger.log("Mise à jour terminée.");
}


function getExistingReservations(sheet) {
  var data = sheet.getDataRange().getValues();
  var header = data.shift(); // Récupération des en-têtes
  var existingReservations = {};

  data.forEach((row, index) => {
    var refReseliva = row[1]; // La colonne 1 contient la référence de réservation
    if (!existingReservations[refReseliva]) {
      existingReservations[refReseliva] = [];
    }
    existingReservations[refReseliva].push({ index: index + 2, row: row }); // Stocker l'index réel dans la feuille (ligne Excel = index+2)
  });

  return { header, existingReservations };
}

function updateExistingRow(sheet, rowIndex, newRow) {
  var range = sheet.getRange(rowIndex, 1, 1, newRow.length);
  range.setValues([newRow]); // Mise à jour avec la nouvelle ligne
}

function deleteRows(sheet, rows) {
  rows.sort((a, b) => b - a); // Trier en ordre inverse pour éviter le décalage des indices
  rows.forEach(row => sheet.deleteRow(row));
}


function traiterMessageOpt (msgTemp, param, sheetError, existingReservations, flagReplace) {
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
      sheetError.appendRow([now,valRefReseliva, error]);
    }
  }
  index = 0;
}

