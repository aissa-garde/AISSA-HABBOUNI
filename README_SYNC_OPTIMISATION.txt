TOX-GARDE — Optimisation de la synchronisation Supabase

Base : Tox-Garde_V71_v77_Correction_Validation_Ancien_Planning(2)

Modification appliquée uniquement à la couche de synchronisation cloud de index.html.

1. Polling Supabase : 5 secondes -> 2 minutes.
2. Synchronisation immédiate au retour sur l'onglet (visibilitychange).
3. Synchronisation immédiate lorsque la fenêtre reprend le focus.
4. Synchronisation immédiate après pageshow.
5. Synchronisation immédiate à la reconnexion Internet (online).
6. Verrou anti-requêtes simultanées pour éviter plusieurs lectures cloud en parallèle.
7. Protection de 15 secondes contre les vérifications événementielles rapprochées.
8. Aucun planning, historique, médecin, score ou autre donnée métier supprimé.
9. Le mécanisme save() reste inchangé : les modifications utilisateur continuent d'être enregistrées localement puis sur Supabase.

Objectif : réduire fortement l'Egress Supabase tout en gardant une actualisation automatique fiable.

Important : cette version conserve le stockage actuel sous forme de snapshot global. Une optimisation ultérieure plus poussée pourrait remplacer la lecture complète par un mécanisme de signal de changement/Realtime côté Supabase, mais cela nécessiterait une adaptation du backend SQL.
