function placerCommande() {
  var stAchat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Achats");
  var stBom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BOM");
  var stVente1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ventes");
  var stVente2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VEC");
  var stTotal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventaire Total");
  var stProduit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Code produits");
  var stOpstock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Operation stock");
  
  var colonne;

  //Valeur de feuille Achat
  var lastRowAchat = stAchat.getLastRow();
  var lastColumnAchat = stAchat.getLastColumn();
  stAchat.getRange(3, 5, lastRowAchat, 1).clear();
  var valeurAchat = stAchat.getSheetValues(3, 1, lastRowAchat, lastColumnAchat);
  var dateAchat = Array.from(Array(valeurAchat.length), () => new Array(1))

  //Valeur de feuille Inventaire Total
  var lastRowTotal = stTotal.getLastRow();
  var lastColumnTotal = stTotal.getLastColumn();
  var valeurTotal = stTotal.getSheetValues(5, 1, lastRowTotal, lastColumnTotal);

  //Valeur de Bom
  var lastRowBom = stBom.getLastRow();
  var lastColumnBom = stBom.getLastColumn();
  var valeurBom = stBom.getSheetValues(2, 1, lastRowBom, lastColumnBom);

  //valeur de Vente
  var aujourd = new Date();


  var lastRowVente1 = stVente1.getLastRow();
  var lastColumnVente1 = stVente1.getLastColumn();
  var valeurVente1 = stVente1.getSheetValues(1, 1, lastRowVente1, lastColumnVente1);
  valeurVente1 = valeurVente1.filter((row, i) =>
      row[0] >= aujourd
    )
    lastRowVente1 = valeurVente1.length;
    lastColumnVente1 = valeurVente1[0].length;

  //Valeur de Code Produit
  var lastRowProduit = stProduit.getLastRow();
  var lastColumnProduit = stProduit.getLastColumn();
  var valeurProduit = stProduit.getSheetValues(2, 1, lastRowProduit, lastColumnProduit);

  //Valeur de Opération stock
  var lastRowOpStock = stOpstock.getLastRow();
  var lastColumnOpStock = stOpstock.getLastColumn();
  var valeurOpStock = stOpstock.getSheetValues(1, 1, lastRowOpStock, lastColumnOpStock);

  var bomTypeColumn = [];
  var leadTime = [];
  var safetyStockTime = [];
  var date = [];

  //Valeur inventaire en main
  var inventaireMain = [];
  var maxRow;

  if (lastRowAchat > lastRowTotal) {
    maxRow = lastRowAchat;
  } else if (lastRowTotal > lastRowProduit) {
    maxRow = lastRowTotal;
  } else {
    maxRow = lastRowProduit;
  }
  var piece;

  //Filtrer pour avoir les valeur d'aujourd.hui et au dessus
  var aujourd = new Date();
  valeurOpStock = valeurOpStock.filter((row, i) =>
    row[0] >= aujourd || i == 0
  )

  var tempStock;

  for (i = 0; i < maxRow; i++) {
    if (i < lastRowTotal) {
      piece = valeurTotal[i][0];
      inventaireMain[piece] = valeurTotal[i][4] - valeurTotal[i][3];
      tempStock = valeurOpStock.filter((row, i) => row[1] == piece);
      for (y = 0; y < tempStock.length; y++) {
        inventaireMain[piece] += tempStock[y][2];
      }
      if (inventaireMain[piece] < 0 && !(piece in date) && piece != "") {
        date[piece] = new Date();
      }
    }

    if (i < lastRowProduit) {
      piece = valeurProduit[i][0];
      leadTime[piece] = valeurProduit[i][2];
      safetyStockTime[piece] = valeurProduit[i][3];
    }

  }


  i = 1;
  while (valeurBom[0][i] != "") {
    bomTypeColumn[valeurBom[0][i]] = i;
    i++;
  }

  var j;
  var k;
  var splitType = [];

  for (i = 0; i < lastRowVente1; i++) {
    splitType = valeurVente1[i][5].split("+");
    for (j = 0; j < splitType.length; j++) {
      k = 1;
      if (splitType[j] in bomTypeColumn) {
        colonne = bomTypeColumn[splitType[j]];
        while (valeurBom[k][0] != "") {
          piece = valeurBom[k][0];
          inventaireMain[piece] -= valeurBom[k][colonne];

          if (inventaireMain[piece] <= safetyStockTime[piece] && !(piece in date) && valeurBom[k][colonne] != safetyStockTime[piece]) {
            date[piece] = valeurVente1[i][0];
          }
          k++;
        }

      }

    }
  }
  var type, item;
  for (const [type, item] of Object.entries(date)) {
    date[type].setDate(date[type].getDate() - leadTime[type]);
  }
  for (i = 0; i < lastRowAchat; i++) {
    type = valeurAchat[i][0];
    if (type in date) {
      dateAchat[i][0] = date[type];
    }
  }
  affichageAchat(dateAchat, stAchat)
}

function affichageAchat(tableauPrinc, stCasier) {
  var height = tableauPrinc.length;
  //Check the max width.
  var width = 0;
  for (var i = 0; i < height; i++) {
    width = Math.max(width, tableauPrinc[i].length);
  }
  //Add the required empty values to convert the jagged array to a 2D array
  var temp;
  for (var i = 0; i < height; i++) {
    temp = tableauPrinc[i].length;
    for (var j = 0; j < width - temp; j++) {
      tableauPrinc[i].push('');
    }
  }
  stCasier.getRange(3, 5, height, width).setValues(tableauPrinc);
}





