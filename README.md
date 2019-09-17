## "Donnees par type de travaux :"
* Dans cette liste vont etre affichés les sommes des depenses previsionnels et effectes par type de traveaux.
      
## "Donnees par entreprise :"
* Dans cette liste vont etre affichés les sommes des depenses previsionnels et effectes par entreprise.
    
## "Donnees d'ensemble :"
### Les informations sont tirés du fichier "GV compta synthese.xlsx"
* Le "Budget total" est fourni manuellement dans le fichier .xlsx, premiere feuil de calcul, celule B1.
* Le "Budget restant previsionel" = budget - depenses effectives - depenses previsionelles
* Le "Budget restant effectif" = budget - depenses effectives
* "Depenes prevues" = somme des toutes les depenses avec status "not started" et "started"
* "Depenses effectives" = somme des toutes les depenses avec status "finished"

## Button "Nouvelle facture"
### Permet d'ajouter une nouvelle ligne "facture".
* "Type de travaux" vas defenir dans quelle fichier excel ira la facture
* "Nom de l'entreprise" vas defenir dans quelle feuille de calcul ira la facture
* Date de debut et de fin : options arbitraires. Pas d'impact sur la logique de logiciel pour l'instant.
* "Etat de traveaux" (options arbitraires)
  * "Not started" et "Started" : le prix va etre considere comme depense previsionelle
  * "Finished" : le prix va etre considere comme efectif
  * "Canceled" : le prix n'est plus consideré pour les calculs
* Commentaires :
#### Champ arbitraire. Juste pour faire la distinction plus facile entre les factures.
* Button "Confirmer": Sauvgarder les informations dans excel
* Annuler: Annuler

## Button "Facture en cours"
### Affiche la liste des toutes les factures existantes. Double click pour editer la facture.
#### Le mode d'edition est similaire a l'ajout de la nouvelle facture, avec la difference de "Type de traveaux" et "Nom de l'entreprise" etant bloque sur les valeurs existants.
