# CHANGELOG — tableau de bord, icône de déconnexion, page « Mon bilan » (09/10/2026)

Aucune règle métier, aucun calcul, aucune donnée modifiés.

- **Tableau de bord** : la cellule « Absents aujourd'hui » est supprimée (2 compteurs restent) ; deux cases ajoutées sous « Disponibilités » : **Congé de maternité** et **Maladie longue durée (MLD)**, visibles par tous les comptes.
- **Icône de déconnexion** : nouvelle icône « power » (anneau + barre) sur ordinateur et mobile ; l'anneau suit la couleur d'en-tête de la palette et la barre la couleur d'accent.
- **Pages non autorisées** : à la connexion / déconnexion, une page non autorisée pour le compte (ex. « Mon bilan » resté ouvert dans le compte admin) est remplacée par le tableau de bord ; l'admin ne peut plus ouvrir « Mon bilan ».
- `sw.js` : cache v41.

# CHANGELOG — corrections planning, menu et icône (09/10/2026)

Aucune règle de points, aucun moteur d'équité, aucune donnée modifiés.

- **Suppression congés/récups (compte Médecin)** : après validation d'un planning, seuls les congés/récups qui chevauchent la **période validée** sont verrouillés (`isAbsenceValidated`). Les autres périodes restent supprimables. L'administrateur peut tout supprimer.
- **Bouton « 🔍 Aperçu »** (page Planning) : aperçu en lecture seule, plein écran, paysage (verrouillage d'orientation sur mobile, rotation CSS en secours). Bouton animé (reflet + loupe), désactivé si « réduire les animations ».
- **Menu médecin** : « Paramètres » n'apparaît plus pour un compte Médecin (règle CSS + garde dans le script).
- **Menu mobile** : fermeture automatique au clic en dehors ; réduction automatique sur tablette / téléphone en paysage ; défilement vertical du menu (téléphone en paysage).
- **Planning** : ligne des dates figée en haut et colonne « Médecin » figée à gauche, en mode normal et plein écran (le tableau défile dans son cadre).
- **Icône / écran de démarrage** : logo recentré avec marges, icônes `any` et `maskable` séparées, `apple-touch-icon` dédiée, `background_color` blanc. Pour voir le changement : désinstaller puis réinstaller l'application.
- `sw.js` : cache v40.

# CHANGELOG — nouveau style des boutons

Présentation uniquement ; aucune règle métier modifiée.

- Boutons arrondis avec soulèvement au survol, retour à l'appui, anneau de focus clavier et état désactivé lisible.
- **Les couleurs suivent la palette réglable** : bouton principal en dégradé léger dans la teinte « secondaire » de la palette,
  boutons secondaires avec fond/bordure/texte de la palette, variante pastel (`btn-soft`) basée sur la couleur « active ».
- Vert fixe pour valider (`btn-ok` : « Valider » des justificatifs, « Accepter » d'un changement) et rouge fixe pour supprimer/refuser (`danger`),
  pour que leur sens ne dépende pas de la palette.
- `sw.js` : cache v39.

# CHANGELOG — nouveau menu

Présentation et organisation du menu uniquement ; aucune règle métier modifiée.

- Menu en carte arrondie, en trois sections : **Principal** (Tableau de bord, Planning, Congés, Génération), **Gestion** (Équipe, Récupération, Bilan, Comptes)
  et **Mon espace** (Mes gardes, Mon bilan) ; pour l'Admin, cette dernière section s'appelle « Suivi » (Historique des changements).
- Pastille d'icône colorée par section ; élément actif teinté avec barre d'accent. Mode réduit conservé (sections masquées, badge sur l'icône).
- **Badge sur « Mes gardes »** (compte Médecin) : nombre de demandes de changement reçues en attente de réponse, dans le menu et dans la barre du bas.
- Mobile : barre du bas flottante et arrondie ; « Mes gardes » y figure pour les comptes Médecin.
- `sw.js` : cache v38.

# CHANGELOG — nouveau style du tableau de bord

Présentation uniquement : mêmes données, aucune règle ni calcul métier modifiés.

- Cellules pastel à coins arrondis avec liseré de couleur en bas, une couleur par type : garde J (bleu), N (violet), G (vert), F (ambre) ;
  congé (rose/bleu), récupération (turquoise), indisponible (orange) ; compteurs en gris-bleu.
- Trois compteurs en haut : Médecins actifs, **Gardes aujourd'hui** et **Absents aujourd'hui** (calculés avec les données déjà chargées).
- Sections « Gardes du jour » et « Disponibilités », chaque disponibilité avec un badge de nombre et la liste des noms.
- Adapté au mobile (2 colonnes pour les gardes, compteurs compacts).
- `sw.js` : cache v37.

# CHANGELOG — « Mon bilan » (compte médecin)

Ajout uniquement : aucune règle de points, aucun calcul du Bilan Admin, aucune donnée modifiés.

- Nouvelle rubrique **Mon bilan** (menu des comptes Médecin, 📊), en lecture seule, avec le même sélecteur d'année que le Bilan Admin
  (année en cours par défaut, option « Toutes les années »).
- Affiche : mes points, ma durée de participation, mes points/mois, et la **moyenne de l'équipe** (points/mois) avec l'écart en %.
  Mêmes chiffres que le Bilan Admin (calcul réutilisé : `bilanComputeYear`).
- Aucun nom de collègue et aucun classement dans cette rubrique.
- La comparaison avec l'équipe n'est affichée qu'à partir de 3 médecins participants (en dessous, la moyenne permettrait de déduire
  les chiffres des autres).
- Correctif : le bouton « Mon bilan » n'apparaît plus dans le menu du compte Administrateur.
- `sw.js` : cache v36.

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
