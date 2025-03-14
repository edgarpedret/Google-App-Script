function enviarInfoEsdevenimentAMembres() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obté les dues fulles pel seu nom
  var fullMembres = ss.getSheetByName("Full Usuaris"); // Fulla1 conté la informació dels membres
  var fullEsdeveniments = ss.getSheetByName("Full Events"); // Fulla2 conté la informació dels esdeveniments

  // Llegeix totes les dades de les dues fulles
  var membres = fullMembres.getDataRange().getValues(); 
  var esdeveniments = fullEsdeveniments.getDataRange().getValues(); 

  // Expressió regular per validar correus electrònics
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  // Recorre tots els membres (comença a la fila 3, que és la index 2)
  for (var i = 2; i < membres.length; i++) { 
    var nom = membres[i][0];  // Columna A - Nom
    var primer_cognom = membres[i][1]; // Columna B - Cognom
    var segon_cognom = membres[i][2]; // Columna C - Segon Cognom
    var localitat = membres[i][3]; // Columna D - Codi Postal
    var email = membres[i][4]; // Columna E - Correu electrònic
    var edat = membres[i][5]; // Columna F - Edat
    var tractament = membres[i][6]; // Columna G - Tractament
    
    // Validació del correu electrònic
    if (email && emailRegex.test(email)) { 
      var missatge = tractament + " " + nom + " " + primer_cognom + " " + segon_cognom + 
                     ", que vius a " + localitat + " i tens " + edat + " anys. ";

      // Afegim la informació dels esdeveniments en un sol paràgraf
      for (var j = 2; j < esdeveniments.length; j++) { 
        var localitat_esdeveniment = esdeveniments[j][0];  // Columna A - Localitat
        var dia = esdeveniments[j][1];        // Columna B - Dia
        var mes = esdeveniments[j][2];        // Columna C - Mes
        var hora = esdeveniments[j][3];       // Columna D - Hora event
        var organitzacio = esdeveniments[j][4]; // Columna E - Qui ho organitza
        var dj = esdeveniments[j][5];         // Columna F - DJ/GRUP

        missatge += "El proper esdeveniment serà a " + localitat_esdeveniment + 
                    " el dia " + dia + " de " + mes + " a les " + hora + 
                    ", organitzat per " + organitzacio + 
                    " i amb l'actuació de " + dj + ". ";
      }

      missatge += "Us hi esperem! Salutacions.";

      var assumpte = "Esdeveniments de l'Associació de Joves a " + localitat;
      
      // Enviar correu
      MailApp.sendEmail(email, assumpte, missatge);
      Logger.log("Correu enviat a: " + email);
    } else {
      Logger.log("Correu electrònic invàlid o buit: " + email + " (Usuari: " + nom + " " + primer_cognom + ")");
    }
  }
}

