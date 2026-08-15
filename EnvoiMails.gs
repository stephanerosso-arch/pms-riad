////// envoyer Mail //////
function envoyerMails () {
  var socios = 'guilmour101@gmail.com, adil.gougou@yahoo.fr,'
  var abd  = 'sales@e-webhotels.com,'
  var ibt = 'ibtissamzouhri@ambreetepices.com,';
  var admin = 'assistanteambreetepices@gmail.com,';
  var kech = 'ambreetepices@gmail.com,'
  var kar = 'Abdelkarim.aithimmi@outlook.com,'

  var email = 'guilmour101@gmail.com,';
  var now = new Date();
  var dateTxt = now.getDate() + '/' +  (now.getMonth()+1)  + '/' + now.getFullYear()
  
  sendEmail (socios+abd+ibt+kech, "Etat des vente nuitée au " + dateTxt , '1jxvJt9GL3dSUsANBr2cFhdCv8tbe2_vNTOXckxxN0Bo', "Synthèse Vente Nuitée", 'C1:I65');
  sendEmail (email, "Etat de la caisse au " + dateTxt , '118QsCu9-jQ87Sgii-IodnHL5-kGV2cK4kH0qwsds3p0', "TdB", 'B1:L19');
  sendEmail (socios+kar+admin, "Dépenses Cash au " + dateTxt , '118QsCu9-jQ87Sgii-IodnHL5-kGV2cK4kH0qwsds3p0', "TdB", 'B7:J19');

}


function envoyerMailErreur() {
  var ss = SpreadsheetApp.openById('1IlKOUxDD77ah4pgb32Obrk7ilqaDdWI0QQofwvHZys8');
  var sheet = ss.getSheetByName("Erreur Resa");
  var dataSyn = sheet.getDataRange().getValues();
  
  var email = 'guilmour101@gmail.com, adil.gougou@yahoo.fr, ambreetepices@gmail.com';
//  var email = 'guilmour101@gmail.com';
  var now = new Date();
  var i = sheet.getLastRow();
  sendEmail (email, (i-1) + " Erreur sur Fichier Resa à corriger au " + now.getDate() + '/' +  (now.getMonth()+1)  + '/' + now.getFullYear() , sheet, 'A1:I' + i);
}

function envoyerMailAlerte(valRefReseliva,valTitulaire,valDateDeb, valNbNuit) {
  var email = 'guilmour101@gmail.com, assistanteambreetepices@gmail.com';
  var subject = "Alerte client premium " + valTitulaire + " pour " + valNbNuit + " Nuitées";
  var body = '<div style="text-align:center;display: inline-block;font-family: arial,sans,sans-serif">'
  body += '<H2>Alerte automatique, client à contacter pour organisation du séjour</H2>';
  body += '<H3>' + "Ref Reseliva : " + valRefReseliva + " Nom client : " + valRefReseliva + " Date arrivée : " + valDateDeb + '</H3>';
  body += '</div>';
  GmailApp.sendEmail(email, subject, "Requires HTML",{htmlBody:body});

}


// code piqué

function sendEmail (recipient, subject, idSs, sheet, cellule) {
  var ss = SpreadsheetApp.openById(idSs)
  var sheet = ss.getSheetByName(sheet);
  var schedRange = sheet.getRange(cellule);
//  var file = DriveApp.getFileById(sheet.getParent().getId())

  // Put Name & Date into email first.
  // We only want the schedule within borders, so
  // these are handled separately.
  var body = '<div style="text-align:center;display: inline-block;font-family: arial,sans,sans-serif">'
  body += '<H1>'+ subject +'</H1>';
  body += '<H2>traitement automatique</H2>';
  body += getHtmlTable(schedRange);
  body += '</div>';
//  GmailApp.sendEmail(recipient, subject, "Requires HTML", {attachments: [file.getAs(MimeType.PDF)], htmlBody:body})
  GmailApp.sendEmail(recipient, subject, "Requires HTML",{htmlBody:body});
  
}

/**
* Return a string containing an HTML table representation
* of the given range, preserving style settings.
*/
function getHtmlTable(range){
  var ss = range.getSheet().getParent();
  var sheet = range.getSheet();
  startRow = range.getRow();
  startCol = range.getColumn();
  lastRow = range.getLastRow();
  lastCol = range.getLastColumn();
  
  // Read table contents
  var data = range.getValues();
  
  // Get css style attributes from range
  var fontColors = range.getFontColors();
  var backgrounds = range.getBackgrounds();
  var fontFamilies = range.getFontFamilies();
  var fontSizes = range.getFontSizes();
  var fontLines = range.getFontLines();
  var fontWeights = range.getFontWeights();
  var horizontalAlignments = range.getHorizontalAlignments();
  var verticalAlignments = range.getVerticalAlignments();
  
  // Get column widths in pixels
  var colWidths = [];
  for (var col=startCol; col<=lastCol; col++) { 
    colWidths.push(sheet.getColumnWidth(col));
  }
  // Get Row heights in pixels
  var rowHeights = [];
  for (var row=startRow; row<=lastRow; row++) { 
    rowHeights.push(sheet.getRowHeight(row));
  }
  
  // Future consideration...
  var numberFormats = range.getNumberFormats();
  
  // Build HTML Table, with inline styling for each cell
  var tableFormat = 'style="border:1.5px solid black;border-collapse:collapse;text-align:center" border = 1.5 cellpadding = 5';
  var html = ['<table '+tableFormat+'>'];
  // Column widths appear outside of table rows
  for (col=0;col<colWidths.length;col++) {
    html.push('<col width="'+colWidths[col]+'">')
  }
  // Populate rows
  for (row=0;row<data.length;row++) {
    html.push('<tr height="'+rowHeights[row]+'">');
    for (col=0;col<data[row].length;col++) {
      // Get formatted data
      var cellText = data[row][col];
      if (cellText instanceof Date) {
        cellText = Utilities.formatDate(
          cellText,
          ss.getSpreadsheetTimeZone(),
          'MMM/d EEE');
      }
      var style = 'style="'
      + 'color: ' + fontColors[row][col]+'; '
      + 'font-family: ' + fontFamilies[row][col]+'; '
      + 'font-size: ' + fontSizes[row][col]+'; '
      + 'font-weight: ' + fontWeights[row][col]+'; '
      + 'background-color: ' + backgrounds[row][col]+'; '
      + 'text-align: ' + horizontalAlignments[row][col]+'; '
      + 'vertical-align: ' + verticalAlignments[row][col]+'; '
      +'"';
      html.push('<td ' + style + '>'
                +cellText
                +'</td>');
    }
    html.push('</tr>');
  }
  html.push('</table>');
  
  return html.join('');
}

