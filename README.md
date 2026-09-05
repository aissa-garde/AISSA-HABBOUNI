# Tox-Garde V50

Version reconstruite à partir de la version stable V37, avec correction ciblée et uniforme de la gestion des dates civiles.

## Correction principale
Les dates saisies manuellement et les dates provenant d'Excel sont interprétées en date locale (année-mois-jour), sans conversion UTC. Cela évite le décalage d'un jour des gardes G/F.

Les autres fonctions de la version stable sont conservées : génération, J/N, G/F, congés/indisponibilités, récupération, historique, confirmation, équité et export.


## Correctifs V71 — droits Médecin
- Un compte Médecin peut enregistrer uniquement ses propres jours de récupération depuis la partie Planning/Garde.
- Un compte Médecin ne peut pas modifier le planning ni le valider.
- Un compte Médecin ne voit que les plannings validés dans son sélecteur de consultation.
- Lorsqu’aucun planning n’existe, la zone de planning reste vide avec un message explicite.
- Les fonctions de génération, équipe, validation, historique et récupération administrateur existantes sont conservées.
