# Tox-Garde — version corrigée (audit du 08/10/2026)

PWA de gestion et de planification des gardes (HTML/JS + Supabase).

## Déploiement
Fichiers à publier : `index.html`, `sw.js`, `manifest.webmanifest`, `icons/`.
Le fichier `supabase_list_accounts.sql` est à exécuter dans Supabase uniquement s'il n'est pas déjà installé.
Le fichier `supabase_get_revision.sql` est à exécuter une fois dans Supabase (SQL Editor) pour activer la vérification légère
de la synchronisation ; sans lui, l'application fonctionne comme avant (téléchargement complet à chaque contrôle).
Après publication, recharger l'application une fois (le service worker passe en `tox-garde-pwa-v39`).

## Actions manuelles OBLIGATOIRES
1. **Changer le mot de passe du compte administrateur initial** dans Supabase : l'ancien mot de passe
   figurait en clair dans les versions précédentes du code (et reste visible dans l'historique GitHub).
2. Demander à tous les utilisateurs de recharger l'application (le contrôle de conflit n'est actif que
   pour les postes qui exécutent cette version).

## Fichiers retirés de l'archive
`app.js`, `test.js`, `_block0.js`, `_block1.js`, `_index_inline_1.js`, `index.html.before_guard_fix`,
icônes en double à la racine, anciens README (contenu fusionné ici).

## Historique résumé
- Nouveau style des boutons, couleurs liées à la palette réglable (voir CHANGELOG.md).
- Nouveau menu : sections Principal / Gestion / Mon espace, badge des demandes de changement sur « Mes gardes » (voir CHANGELOG.md).
- Nouveau style du tableau de bord : cellules pastel avec liseré de couleur, compteurs gardes/absents du jour (voir CHANGELOG.md).
- Rubrique **Mon bilan** (compte Médecin) : mes points, ma durée, mes points/mois et moyenne anonyme de l'équipe (voir CHANGELOG.md).
- Synchronisation allégée : pas de téléchargement en arrière-plan, déconnexion après 15 min d'inactivité, contrôle toutes les 5 min, vérification légère par révision (voir CHANGELOG.md).
- Rubrique **Bilan** (Admin) : points/mois par année, graphiques, exports PDF/Excel (voir CHANGELOG.md).
- V18.4 : menu repliable, justificatifs via Supabase, correctifs login et échanges G/F.
- V18.5 : date d'intégration, équité normalisée, protection contre l'écrasement par une copie plus ancienne.
- Générations de planning : numéro de génération par mois, sélecteur de génération.
- Audit 08/10/2026 : voir CHANGELOG.md.
