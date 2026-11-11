-- Chat turns in the ChatSlide editor since :since, one row per user. Prod, read-only:
-- psql <prod> -v since="'2026-11-10'" -f /Users/quanlaili/proj/PPTist/scripts/editor-turns.sql
with t as (
  select id, uid, timestamp, payload::jsonb p
  from user_action_logs
  where endpoint = 'editor_turn' and timestamp > :since
)
select t.uid, us.subscription_tier,
  (select country_code from user_activity x where x.uid = t.uid
     order by created_at desc limit 1) country,
  json_agg(json_build_object('id', t.id, 'ts', t.timestamp,
    'proj', t.p->>'project', 'msg', left(t.p->>'message', 300),
    'reply', left(t.p->>'reply', 300), 'ok', t.p->'ok', 'steps', t.p->'steps')
    order by t.timestamp) turns
from t left join users us on us.uid = t.uid
group by 1, 2, 3;
