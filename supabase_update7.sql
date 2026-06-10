-- Execute no Supabase > SQL Editor > New Query
-- Cria tabela de configurações por usuário

create table if not exists user_config (
  id        bigint generated always as identity primary key,
  user_id   uuid references auth.users(id) unique,
  loja1     text default 'Loja 1',
  loja2     text default 'Loja 2',
  criado_em timestamptz default now()
);

alter table user_config enable row level security;

create policy "config_select" on user_config for select using (auth.uid() = user_id);
create policy "config_insert" on user_config for insert with check (auth.uid() = user_id);
create policy "config_update" on user_config for update using (auth.uid() = user_id);

-- Insere configuração padrão para usuário existente (Gran Shop e Aishael Mix)
insert into user_config (user_id, loja1, loja2)
values ('3ccdd460-3798-4d87-9571-33c16be7071f', 'Gran Shop', 'Aishael Mix')
on conflict (user_id) do nothing;
