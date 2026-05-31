-- Execute no Supabase > SQL Editor > New Query

-- ── Corrige constraints de pedidos ───────────────────────────────────────────
ALTER TABLE pedidos DROP CONSTRAINT IF EXISTS pedidos_user_pedido;
ALTER TABLE pedidos DROP CONSTRAINT IF EXISTS pedidos_user_id_plataforma;
ALTER TABLE pedidos DROP CONSTRAINT IF EXISTS pedidos_id_plataforma_key;

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'pedidos_user_pedido'
  ) THEN
    ALTER TABLE pedidos ADD CONSTRAINT pedidos_user_pedido UNIQUE (user_id, id_pedido);
  END IF;
END $$;

-- Colunas que podem estar faltando
ALTER TABLE pedidos ADD COLUMN IF NOT EXISTS status_interno text DEFAULT '';
ALTER TABLE pedidos ADD COLUMN IF NOT EXISTS nota_revisao  text DEFAULT '';
ALTER TABLE pedidos ADD COLUMN IF NOT EXISTS loja          text DEFAULT '';
ALTER TABLE pedidos ADD COLUMN IF NOT EXISTS preco         numeric(12,2) DEFAULT 0;

-- ── Corrige tabela devoluções ─────────────────────────────────────────────────
-- Remove policies antigas (sem erro se não existirem)
DROP POLICY IF EXISTS "dev_leitura"     ON devolucoes;
DROP POLICY IF EXISTS "dev_escrita"     ON devolucoes;
DROP POLICY IF EXISTS "dev_atualizacao" ON devolucoes;
DROP POLICY IF EXISTS "dev_exclusao"    ON devolucoes;

-- Recria policies corretamente
ALTER TABLE devolucoes ENABLE ROW LEVEL SECURITY;

CREATE POLICY "dev_leitura"     ON devolucoes FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "dev_escrita"     ON devolucoes FOR INSERT WITH CHECK (auth.uid() = user_id);
CREATE POLICY "dev_atualizacao" ON devolucoes FOR UPDATE USING (auth.uid() = user_id);
CREATE POLICY "dev_exclusao"    ON devolucoes FOR DELETE USING (auth.uid() = user_id);

-- Corrige constraint de devoluções
ALTER TABLE devolucoes DROP CONSTRAINT IF EXISTS devolucoes_user_pedido;

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'devolucoes_user_pedido'
  ) THEN
    ALTER TABLE devolucoes ADD CONSTRAINT devolucoes_user_pedido UNIQUE (user_id, id_pedido);
  END IF;
END $$;
