# ToxGarde — corrections intégration / équité / synchronisation

Base utilisée : ToxGarde_V18_5_AUTO_SYNC vfffffffff.

Modifications ciblées dans `index.html` :
- Ajout de `participationStart` / Date intégration dans l'équipe.
- Nouveau médecin actif : date du jour automatiquement.
- Inactif -> Actif : date du jour seulement si aucune date n'existe.
- Générateur : aucune nouvelle garde J/N avant la date d'intégration.
- Équité : charge normalisée par les jours de participation effective.
- Maladie longue durée et congé de maternité : exclus du dénominateur d'équité.
- Congé / indisponibilité / récup : logique existante conservée.
- Ajout des deux types d'absence comme valeurs de données.
- Import Excel : colonne `Date intégration` prise en charge.
- Protection supplémentaire contre l'écrasement local par une copie Supabase plus ancienne.
- Générateur de planning et règles de repos 48 h conservés.

Vérifications :
- Syntaxe JavaScript : OK.
- Contrôles statiques ciblés : OK.


## Gestion des générations de planning — 2026-10-06
- Ajout d'un numéro de génération par mois (Génération 1, 2, 3...).
- Une seule génération, conservée sous la clé courante du mois, est marquée ACTUELLE.
- Les générations précédentes restent consultables et sont marquées Ancienne.
- Ajout d'un sélecteur de génération directement dans la navigation du planning.
- Conservation des anciennes structures historiques pour compatibilité avec les données existantes.
- Aucun changement volontaire du moteur d'équité, des règles 48 h, G/F, J/N, Jr ou AUTO SYNC.
