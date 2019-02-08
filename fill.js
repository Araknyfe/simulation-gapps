function fillBc(doc) {
    var body = doc.getBody();
    var header = doc.getHeader();

    //PRESIDENT

    if(president.sexe == 'M') {
        body.replaceText('{je_president_titre}', 'Mr');
        body.replaceText('{je_president_fonction}', '');
    } else if(president.sexe == 'F') {
        body.replaceText('{je_president_titre}', 'Mme');
        body.replaceText('{je_president_fonction}', 'e');
    }

    body.replaceText('{je_president_prenom}', president.prenom);
    body.replaceText('{je_president_nom}', president.nom);
    body.replaceText('{je_president_portable}', president.portable);

    // ETUDE
    body.replaceText('{etude_titre}', etude.nom);
    body.replaceText('{date_debut_simple}', etude.dateDebut);
    body.replaceText('{num_CC}', etude.cc); 
    body.replaceText('{etude_duree_semaine}', etude.nbSemaine); 
    body.replaceText('{etude_nombre_JEH}', etude.nbJEH); 

    //PRIX
    body.replaceText('{etude_montant_HT}', prix.montantHT); 
    body.replaceText('{frais_HT}', prix.fraisDossier); 
    body.replaceText('{total_HT}', prix.totalHT); 
    body.replaceText('{total_TVA}', prix.totalTVA); 
    body.replaceText('{total_TTC}', prix.totalTTC); 
    body.replaceText('{etude_acompte_HT}', prix.acompteHT); 
    body.replaceText('{etude_acompte_pourcentage}', prix.acomptePct); 
    body.replaceText('{taux_tva}', '20'); 
    body.replaceText('{etude_reste_HT}', prix.resteHT); 
    body.replaceText('{etude_reste_TTC}', prix.resteTTC); 


    //DATE CREATION

    body.replaceText('{general_date_creation_simple}', getCreationDate()); 


    // RESPO COMMERCIAL
    body.replaceText('{auteur_prenom}', responsableCommercial.prenom);
    body.replaceText('{auteur_nom}', responsableCommercial.nom);
    body.replaceText('{auteur_portable}', responsableCommercial.portable);
    body.replaceText('{auteur_email}', responsableCommercial.email);


    // CLIENT 
    if(client.sexe == 'M') {
        body.replaceText('{client_titre}', 'Mr');
        body.replaceText('{client_fonction}', '');
    } else if(client.sexe == 'F') {
        body.replaceText('{client_titre}', 'Mme');
        body.replaceText('{client_fonction}', 'e');
    }
    body.replaceText('{client_societe}', etude.clientSociete); 
    body.replaceText('{client_prenom}', client.prenom); 
    body.replaceText('{client_societe}', client.nom); 
    body.replaceText('{client_portable}', client.portable);
    body.replaceText('{client_email}', client.email);
    body.replaceText('{client_adresse}', client.adresse);
    body.replaceText('{client_code_postal}', client.codePostal);
    body.replaceText('{client_ville}', client.ville);
}

function fillURL(url) {
    var sheetActive = SpreadsheetApp.getActiveSheet();

    sheetActive.getRange("F1").setValue(url);
}