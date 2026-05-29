-- Execute no Supabase > SQL Editor > New Query

-- Recria tabela pedidos com novo schema
drop table if exists pedidos cascade;

create table pedidos (
  id             bigint generated always as identity primary key,
  user_id        uuid references auth.users(id),
  id_pedido      text not null,
  status_pedido  text default '',
  data_envio     text,
  hora_pagamento text,
  produto        text default '',
  preco          numeric(12,2) default 0,
  quantidade     text default '',
  variacao       text default '',
  destinatario   text default '',
  loja           text default '',
  status_interno text default '',
  nota_revisao   text default '',
  criado_em      timestamptz default now(),
  constraint pedidos_user_pedido unique (user_id, id_pedido)
);

alter table pedidos enable row level security;
create policy "leitura_por_usuario"    on pedidos for select using (auth.uid() = user_id);
create policy "escrita_por_usuario"    on pedidos for insert with check (auth.uid() = user_id);
create policy "atualizacao_por_usuario" on pedidos for update using (auth.uid() = user_id);
create policy "exclusao_por_usuario"   on pedidos for delete using (auth.uid() = user_id);

-- Tabela de faturamento mensal
create table if not exists faturamento (
  id        bigint generated always as identity primary key,
  user_id   uuid references auth.users(id),
  mes       text not null,
  loja      text not null,
  valor     numeric(12,2) default 0,
  criado_em timestamptz default now()
);

alter table faturamento enable row level security;
create policy "fat_leitura"    on faturamento for select using (auth.uid() = user_id);
create policy "fat_escrita"    on faturamento for insert with check (auth.uid() = user_id);
create policy "fat_atualizacao" on faturamento for update using (auth.uid() = user_id);
