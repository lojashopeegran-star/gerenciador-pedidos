-- Execute no Supabase > SQL Editor > New Query
-- Aumenta o tamanho do campo status_pedido para aceitar textos longos

ALTER TABLE pedidos ALTER COLUMN status_pedido TYPE text;
ALTER TABLE pedidos ALTER COLUMN produto       TYPE text;
ALTER TABLE pedidos ALTER COLUMN variacao      TYPE text;
ALTER TABLE pedidos ALTER COLUMN destinatario  TYPE text;
ALTER TABLE pedidos ALTER COLUMN notas         TYPE text;

-- Adiciona coluna notas se não existir
ALTER TABLE pedidos ADD COLUMN IF NOT EXISTS notas text DEFAULT '';

-- Limpa pedidos com status inválido para recarregar corretamente
-- (opcional - só execute se quiser limpar os dados antigos com status "A" errado)
-- DELETE FROM pedidos WHERE status_pedido IN ('A', 'Não pago') AND length(status_pedido) < 10;
