-- Tox-Garde : vérification légère de la révision des données.
-- À exécuter UNE FOIS dans Supabase (SQL Editor). Sans ce fichier, l'application
-- fonctionne comme avant (téléchargement complet à chaque contrôle).
--
-- La fonction renvoie seulement {"rev": <numéro de révision>, "updatedAt": <date>}
-- (quelques octets) au lieu de toutes les données. Elle s'appuie sur la fonction
-- existante toxgarde_get_data : la vérification du jeton de session est donc
-- exactement la même que pour la lecture normale des données.
--
-- Hypothèse : toxgarde_get_data(p_token text) renvoie les données sous forme de
-- JSON (json ou jsonb), avec les champs cloudRevision et cloudUpdatedAt.

create or replace function public.toxgarde_get_revision(p_token text)
returns jsonb
language plpgsql
security definer
set search_path = public, extensions
as $$
declare
  d jsonb;
begin
  d := (public.toxgarde_get_data(p_token))::text::jsonb;

  return jsonb_build_object(
    'rev',       coalesce((d->>'cloudRevision')::numeric, 0),
    'updatedAt', coalesce((d->>'cloudUpdatedAt')::numeric, 0)
  );
end;
$$;

grant execute on function public.toxgarde_get_revision(text)
to anon, authenticated;
