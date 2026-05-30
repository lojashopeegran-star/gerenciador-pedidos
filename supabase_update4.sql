-- Execute no Supabase > SQL Editor > New Query

-- Update pedidos table: upsert should UPDATE existing rows (remove ignoreDuplicates)
-- Add devolucoes table

create table if not exists devolucoes (
  id              bigint generated always as identity primary key,
  user_id         uuid references auth.users(id),
  id_pedido       text not null,
  data_criacao    text,
  produto         text,
  variacao        text,
  preco_unidade   numeric(12,2) default 0,
  status_devolucao text,
  quantidade      text,
  motivo          text,
  observacoes     text,
  loja            text,
  criado_em       timestamptz default now(),
  constraint devolucoes_user_pedido unique (user_id, id_pedido)
);

alter table devolucoes enable row level security;
create policy "dev_leitura"    on devolucoes for select using (auth.uid() = user_id);
create policy "dev_escrita"    on devolucoes for insert with check (auth.uid() = user_id);
create policy "dev_atualizacao" on devolucoes for update using (auth.uid() = user_id);
create policy "dev_exclusao"   on devolucoes for delete using (auth.uid() = user_id);
