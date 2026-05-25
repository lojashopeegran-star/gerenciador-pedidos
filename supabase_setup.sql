-- Execute este SQL no Supabase:
-- Painel > SQL Editor > New Query > cole e clique em Run

create table if not exists pedidos (
  id              bigint generated always as identity primary key,
  id_plataforma   text unique not null,
  subpedido       text,
  numero_pedido   text,
  plataforma      text,
  loja            text,
  estado          text,
  hora_pedido     text,
  hora_pagamento  text,
  prazo_envio     text,
  notas           text,
  produto         text,
  variacao        text,
  quantidade      text,
  unidade         text,
  nome_comprador  text,
  id_comprador    text,
  destinatario    text,
  criado_em       timestamptz default now()
);

-- Permite leitura e escrita pública (ajuste se quiser autenticação)
alter table pedidos enable row level security;

create policy "Acesso público" on pedidos
  for all using (true) with check (true);
