function recupDataMailIA() {
  Logger.log("Lancement V4 Recup Mail");
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

  var threads = GmailApp.search('label:testIA');
//  var threads = GmailApp.search('6667003');
  var msgs = GmailApp.getMessagesForThreads(threads);

  for (i = msgs.length -1; i > -1; i--) {
    const objResa = callMistralAI (msgs[i][0].getPlainBody());
    objResa.DateModif = msgs[i][0].getDate()
    console.log(objResa);
    if (objResa.Statut != "Erreur") traiterResa(objResa, param,"dataMail2");
  }

   for (i = threads.length -1; i > -1; i--) {
    var label = GmailApp.getUserLabelByName("saveIA");
    var label2 = GmailApp.getUserLabelByName("testIA");
    label.addToThread(threads[i]);
    label2.removeFromThread(threads[i]);
  }
}

function traiterResa(objResa,param,sheetName) {
  const correspondance   = param.correspondance   ;
  const tabTypeCh   = param.tabTypeCh   ;
  const tabTypeRate = param.tabTypeRate ;
  const tabTypePromo= param.tabTypePromo;
  const tabTypeSup  = param.tabTypeSup  ;
  const tabTypeCom  = param.tabTypeCom  ;
  const tabPeriode  = param.tabPeriode  ;
  const tabContrat  = param.tabContrat  ;

  const sheet = ss.getSheetByName(sheetName);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const refResa = objResa.RefResa;
  let rowIndex = -1;

  // Recherche de la ligne correspondant à la référence de réservation
  values.forEach((r,i) => rowIndex = (r[1] == refResa) ? i+1 : rowIndex) 

  // Si la réservation est trouvée et que le statut est "Annulation"
  if (rowIndex !== -1 && objResa.Statut === "Annulation") {
    // Mise à jour du statut (colonne 3, index 2)
    sheet.getRange(rowIndex, 3).setValue("Annulation");

    // Réinitialisation des cellules des colonnes 15 à 21 (index 14 à 20)
    sheet.getRange(rowIndex, 15, 1, 7).setValue(0);

    // Ajout du texte "Annulé le [date]" à la date de réservation (colonne 4, index 3)
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const originalDate = sheet.getRange(rowIndex, 4).getValue();
    const newDateValue = `Annulé le ${formattedDate} - ${originalDate}`;
    sheet.getRange(rowIndex, 21).setValue(`Annulé le ` + objResa.DateModif);
  }

  // Ajout des nouvelles lignes si statut est "Confirmation" ou "Modification"
  if (((objResa.Statut === "Confirmation") && (rowIndex === -1)) || (objResa.Statut === "Modification")) {
    // Si le statut est "Modification", identifier et supprimer les lignes correspondantes
    if (objResa.Statut === "Modification") {
      const rowsToDelete = [];

      // Identification des lignes à supprimer
      for (let i = 0; i < values.length; i++) {
        if (values[i][1] == refResa) { 
          rowsToDelete.push(i + 1); // Les index de ligne commencent à 1 dans Google Sheets
        }
      }

      // Suppression des lignes en commençant par la fin pour éviter les décalages d'index
      for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        sheet.deleteRow(rowsToDelete[i]);
      }
    }
    

    // Ajout des informations des chambres
    objResa.Rooms.forEach(room => {
      // Il faut reconsituer le tarif avec des règles de gestion dépendant de la plateforme et du type de contrat
      const typeContrat = searchVPartiel(room.TypeContrat,tabTypeRate,4);
      const comContrat = searchVPartiel(objResa.Source,tabTypeCom)
      const supValue = searchV(room.TypeRate,tabTypeSup)

      if (typeContrat == "B2C") room.TotalRoom = room.TotalRoom - supValue*objResa.NbNight*room.NbPax // Les suppléments ne sont pas affichés dans le prix total

      const roomRow = [
        objResa.DateResa,
        objResa.RefResa,
        objResa.Statut,
        objResa.Source,
        objResa.RefSource,
        objResa.Name,
        objResa.Arrival,
        objResa.Departure,
        objResa.NbNight,
        room.RoomType,
        room.TypeRate,
        room.TypeContrat,
        room.NbPax,
        0,
        room.TotalRoom/objResa.NbNight,
        room.TotalRoom/objResa.NbNight*(1-comContrat), // Nuitée net
        supValue*room.NbPax,
        room.TotalRoom*(1-comContrat), // Total Nuitée net
        supValue*objResa.NbNight*room.NbPax,
        objResa.TotalMail,
        "Contrat : " + room.TypeContrat,
        objResa.DateModif
      ];
      sheet.appendRow(roomRow);
    });
  }
}


function analyseEmail() {
  var threads = GmailApp.search('6357192');
  var msgs = GmailApp.getMessagesForThreads(threads);
//  console.log(msgs[0][0].getPlainBody())
//  const objResa = callOpenAI(msgs[0][0].getPlainBody())
  const objResa = callMistralAI(msgs[0][0].getPlainBody())

//  console.log (objResa);
  return objResa;

}

function callMistralAI(bodyText) {
  var apiKey ="dOTB84vNpDcbNjmIjFjiVWijf1BwQiY6" // Mistral
  var url = "https://api.mistral.ai/v1/chat/completions"; // URL de l'API Mistral AI

  var prompt = `
    Voici un e-mail de réservation d'hôtel. Extrait les informations suivantes :
    - Source : Source de réservation (BookingCom, Expedia), prends bien le libellé complet, s'il n'y en a pas c'est "Reseliva" si le champs 'titulaire de la Réservation" est "Riad Ambre & Epices" alors c'est "Reseliva Pannel"
    - Statut : Statut Resa : (Confirmation, Annulation, Modification)
    - RefResa : Référence de réservation
    - RefSource : Référence de la reservation de la source (généralement à coté de la source), si tu ne la trouves pas, tu mets la même que RefResa
    - DateResa : Date de la reservation au format DD/MM/YYY
    - Name : Nom du client
    - Arrival : Date d'arrivée au format DD/MM/YYY
    - Departure : Date de départ au format DD/MM/YYY
    - NbNight : Nombre de nuits
    - NbRoom : Nombre de chambres
    - TotalMail : Total de la réservation (sous forme de nombre)
    puis pour chaque chambre (Rooms) :
    - RoomType : Type de chambre sous la forme ST (Chambre Standard) / DE (Chambre Deluxe)/ SU (Suite Standard) / SX (Suite Deluxe) / SF (Suite familiale), si tu as un doute sur le type de chambre, tu mets "Erreur" dans ce champs
    - TypeRate : Type de Tarif/Service BB/HALF/FULL (Bed & Breakfast, Half-Board, Full Board) généralement situé sous le Type de Chambre
    - TypeContrat : Nom du contrat ou du type de tarif généralement situé sous le Type de Chambre également
    - NbPax : Nb de personnes pour cette chambre
    - TotalRoom : Prix total (sous forme de nombre)

    Email :
    """ ${bodyText} """

    Fournis la réponse en JSON brut, sans formatage Markdown ni balises de code.
    Parfois le mail ne contient pas de réservation, c'est juste un message, tu sauras le reconnaitre car tu ne récupéreras quasiment aucune des données au dessus, dans ce cas tu envoie juste le JSOM avec Statut : "Erreur"
    `;

  var options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      model: "mistral-large-latest", // Choisissez le modèle approprié
      messages: [
        { role: "system", content: "Tu es un assistant spécialisé en extraction de données de réservations hôtelières." },
        { role: "user", content: prompt }
      ],
      temperature: 0.1
    })
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    return JSON.parse(json.choices[0].message.content); // Retourne l'objet JSON extrait
  } catch (e) {
    Logger.log("Erreur API Mistral AI : " + e.toString());
    return null;
  }
}


function callOpenAI(bodyText) {
  var apiKey = "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"; // Remplace par ta clé API OpenAI

  var url = "https://api.openai.com/v1/chat/completions"; // URL open AI

  var prompt = `
    Voici un e-mail de réservation d'hôtel. Extrait les informations suivantes :
    - Source : Source de réservation (Booking, Expedia, etc.)
    - Statut : Statut Resa : (Confirmation, Annulation, Modification)
    - RefResa : Référence de réservation
    - Name : Nom du client
    - Arrival : Date d'arrivée
    - Departure : Date de départ
    - NbNight : Nombre de nuits
    - NbRoom : Nombre de chambres
    - TotalMail : Total de la réservation
    puis pour chaque chambre :
    - RoomType : Type de chambre
    - TypeRate : Type de Tarif/Service (Bed & Breakfast, Half-Board) généralement situé sous le Type de Chambre
    - TotalRoom : Prix total

    Email :
    """ ${bodyText} """

    Fournis la réponse en JSON brut, sans formatage Markdown ni balises de code.
  `;


  var options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      model: "gpt-4o-mini",
      messages: [{role: "system", content: "Tu es un assistant spécialisé en extraction de données de réservations hôtelières."}, {role: "user", content: prompt}],
      temperature: 0.3
    })
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    return JSON.parse(json.choices[0].message.content); // Retourne l'objet JSON extrait
  } catch (e) {
    Logger.log("Erreur API OpenAI : " + e.toString());
    return null;
  }
}


