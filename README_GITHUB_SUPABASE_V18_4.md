# Tox-Garde — GitHub + Supabase V18.4

Cette version part de la version GitHub/Supabase fonctionnelle et y intègre l'interface V18.3.

Conservé :
- `toxgarde_accounts`
- `toxgarde_data`
- `toxgarde_sessions`
- la configuration Supabase existante

Ajouté :
- menu vertical repliable
- logo de connexion
- progression dynamique à la connexion
- gestion des justificatifs via Supabase
- `toxgarde_justifications` et ses RPC

La connexion Supabase existante est conservée dans `index.html`.


## Correctifs intégrés
- Barre de progression dynamique et étapes de connexion Supabase dans `index.html`.
- Après acceptation d'un échange G/F, synchronisation de la garde dans `state.fixed` afin que le tableau week-end / jours fériés affiche immédiatement le nouveau médecin.
