function fillRm(doc) {
    var body = doc.getBody();
    var header = doc.getHeader();

    // SYNERG
    body.replaceText('{je_ape}', ''); //TODO empty

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
    body.replaceText('{etude_sujet}', etude.sujet); 
    body.replaceText('{num_CE}', etude.ce); 

    // RESPO PROJET
    body.replaceText('{auteur_prenom}', responsableProjet.prenom);
    body.replaceText('{auteur_nom}', responsableProjet.nom);
    body.replaceText('{auteur_portable}', responsableProjet.portable);
    body.replaceText('{auteur_email}', responsableProjet.email);

    // ETUDIANT 
    body.replaceText('{num_avenant_etudiant}', rm.name);
    body.replaceText('{etudiant_prenom}', etudiant.prenom);
    body.replaceText('{etudiant_nom}', etudiant.nom);
    body.replaceText('{etudiant_pronom}', etudiant.pronom);

    header.replaceText('{etudiant_prenom}', etudiant.prenom);
    header.replaceText('{etudiant_nom}', etudiant.nom);

    body.replaceText('{etudiant_adresse}', etudiant.adresse);
    body.replaceText('{etudiant_code_postal}', etudiant.codePostal);
    body.replaceText('{etudiant_ville}', etudiant.ville);
    body.replaceText('{etudiant_portable}', etudiant.portable);
    body.replaceText('{etudiant_mail}', etudiant.mail);

    body.replaceText('{etudiant_poste}', etudiant.poste);
    body.replaceText('{etudiant_missions}', etudiant.descriptionMissions);

    if(etudiant.sexe == 'M') {
        body.replaceText('{etudiant_e}', '');
    } else if(etudiant.sexe == 'F') {
        body.replaceText('{etudiant_e}', 'e');
    }

    body.replaceText('{etudiant_date_fin_mission}', etude.semaineFin);
    body.replaceText('{etudiant_remuneration_th_unit}', etudiant.montantJEH); 
    body.replaceText('{etudiant_nb_JEH}', etudiant.nbJEH); 
    body.replaceText('{etudiant_remuneration_th}', etudiant.remuneration); 

    if(etudiant.nbJEH > 1) {
        body.replaceText('{etudiant_nb_JEH_s}', 's');
    } else {
        body.replaceText('{etudiant_nb_JEH_s}', '');
    }

    // CLIENT 
    body.replaceText('{client_societe}', etude.clientSociete); 
}

function fillURL(url) {
    var sheetActive = SpreadsheetApp.getActiveSheet();

    sheetActive.getRange("F1").setValue(url);
}