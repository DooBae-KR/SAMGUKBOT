-- 시즌 종료 시 지난 시즌의 인사 기록을 시즌 번호별로 보관하는 표 (Supabase SQL Editor에서 한 번 실행)
create table if not exists public.sam_user_history (
  id               bigserial primary key,
  guild_id         bigint      not null,
  season           integer     not null,                 -- 보관한 시즌 번호
  kind             text        not null,                 -- stat(날짜별 기록) | left(이탈자) | alt(부캐 연결)
  cid              text        not null,
  nick             text,
  job              text,
  snap_date        date        not null,                 -- stat: 기록일, left: 이탈일, alt: 보관일
  title            text,
  prosperity       numeric,
  war_merit        numeric,                              -- 주간 무훈
  contribution     numeric,                              -- 주간 공헌
  siege_count      numeric,
  total_merit      numeric,                              -- 누적 무훈
  total_contribution numeric,                            -- 누적 공헌
  camp             text,
  camp_x           integer,
  camp_y           integer,
  main_cid         text,                                 -- alt: 본캐 CID
  extra            jsonb,                                -- left: last_stat 등
  archived_at      timestamptz not null default now(),
  unique (guild_id, season, kind, cid, snap_date)
);
create index if not exists sam_user_history_season_idx on public.sam_user_history (guild_id, season, cid);
alter table public.sam_user_history enable row level security;   -- 봇은 service key로만 접근
