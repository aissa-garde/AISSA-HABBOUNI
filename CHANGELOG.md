# CHANGELOG — synchronisation allégée (consommation Supabase)

Aucune modification des règles métier, du moteur d'équité, des plannings ni de la structure des données.

- **Onglet caché : aucun téléchargement.** Les contrôles périodiques et les signaux temps réel reçus pendant que l'onglet
  est en arrière-plan sont mis en attente ; une seule vérification a lieu au retour sur l'onglet.
- **Déconnexion automatique après 15 minutes d'inactivité**, avec un avertissement (« Rester connecté ») 1 minute avant.
  La déconnexion arrête la synchronisation et le canal temps réel. Elle est différée si une sauvegarde est en cours.
  Si le délai est dépassé pendant que l'onglet était en arrière-plan, la déconnexion a lieu dès le retour sur l'onglet.
- **Contrôle périodique toutes les 5 minutes** (au lieu de 2).
- **Vérification légère avant tout téléchargement** : nouvelle fonction SQL `toxgarde_get_revision`
  (fichier `supabase_get_revision.sql`). Les données ne sont téléchargées que si le numéro de révision distant est plus récent.
  Le contrôle de conflit avant chaque sauvegarde utilise aussi cette vérification. Si la fonction n'est pas installée,
  l'application garde l'ancien fonctionnement (téléchargement complet).
- `sw.js` : cache v34.

# CHANGELOG — rubrique « Bilan » (Admin)

Ajout uniquement : aucune règle de points, aucun générateur, aucune donnée ni fonction existante modifiés.

- Nouvelle rubrique **Bilan** (menu Admin, 📊) : indicateur principal **points / durée réelle de participation**,
  calculé séparément pour chaque année (sélecteur d'année, année en cours par défaut, option « Toutes les années »).
- Points : `recoveryHistoryEntries()` (plannings validés, mêmes gardes que « Récupération ») × `guardPoints()` (J=1, N=2, G=3, F=3).
- Durée : `participationStart` (date d'intégration) et statut `active` existants ; maladie longue durée et congé de maternité
  déduits (même règle que le moteur d'équité) ; précision au jour près, sans arrondi avant le ratio.
  Aucune date de départ n'existant dans les données, la fin d'un médecin inactif ou retiré de l'équipe est estimée
  d'après sa dernière garde validée (signalée par « * »).
- Année en cours : période observée = jusqu'à la dernière date couverte par un planning validé.
- Résumé, 3 graphiques (points, durée, points/mois), tableau triable, exports PDF et Excel (Admin), responsive.
- `sw.js` : cache v33.

# CHANGELOG — correctifs d'audit (08/10/2026)

Aucune modification volontaire du moteur d'équité, des règles de repos 48 h, des gardes G/F, J/N, Jr
ni de la structure des données.

## Sécurité
- Suppression du mot de passe administrateur en clair (`INITIAL_ADMIN`) ; seul l'identifiant réservé reste.
- Ouverture des justificatifs : seuls PDF/PNG/JPEG s'ouvrent ; HTML/SVG refusés (ils pouvaient s'exécuter
  dans l'origine de l'application et lire le token de session).
- Noms de médecins nettoyés à la création/import (`< > " \` \\`) et échappés à l'affichage
  (statistiques, listes de comptes, menus déroulants) ; bouton « Supprimer compte » sans injection possible.
- Ajout d'une Content-Security-Policy (scripts : application + SheetJS/jsPDF ; réseau : Supabase uniquement).

## Fiabilité
- Verrou `cloudDraftLock` libéré dans un `finally` (la synchronisation ne reste plus bloquée après un échec) ;
  nouvelle tentative automatique de sauvegarde du brouillon après 30 s.
- Contrôle de révision (`cloudRevision`) : une sauvegarde est refusée si un autre utilisateur a enregistré
  entre-temps ; les données sont rechargées et l'utilisateur est prévenu.
- Une seule alerte « synchronisation échouée » tant que l'échec persiste.
- Comparaison de dates de la suppression du planning actuel en heure locale (`iso()`), plus de décalage UTC.

## Synchronisation
- Polling de secours 5 s → 2 min (Realtime + focus/visibilité/online conservés).

## Maintenance
- `sw.js` : v32, `app.js` retiré de la liste de cache.
- Archive nettoyée, README unifié.
