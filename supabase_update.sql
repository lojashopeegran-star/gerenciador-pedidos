-- Execute no Supabase > SQL Editor > New Query

-- Adiciona coluna user_id na tabela pedidos
alter table pedidos add column if not exists user_id uuid references auth.users(id);

-- Remove constraint antiga e cria nova com user_id
alter table pedidos drop constraint if exists pedidos_id_plataforma_key;
alter table pedidos add constraint pedidos_user_id_plataforma unique (user_id, id_plataforma);

-- Atualiza policies para filtrar por usuário logado
drop policy if exists "leitura_publica" on pedidos;
drop policy if exists "escrita_publica" on pedidos;
drop policy if exists "atualizacao_publica" on pedidos;
drop policy if exists "exclusao_publica" on pedidos;

create policy "leitura_por_usuario" on pedidos
  for select using (auth.uid() = user_id);

create policy "escrita_por_usuario" on pedidos
  for insert with check (auth.uid() = user_id);

create policy "atualizacao_por_usuario" on pedidos
  for update using (auth.uid() = user_id);

create policy "exclusao_por_usuario" on pedidos
  for delete using (auth.uid() = user_id);
