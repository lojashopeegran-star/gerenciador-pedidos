-- Execute no Supabase > SQL Editor > New Query
-- Cole e execute este SQL para verificar e corrigir tudo

-- 1. Verifica se há pedidos sem user_id
SELECT COUNT(*) as pedidos_sem_user_id FROM pedidos WHERE user_id IS NULL;

-- 2. Verifica as policies atuais de pedidos
SELECT policyname, cmd FROM pg_policies WHERE tablename = 'pedidos';

-- 3. Recria todas as policies de pedidos do zero
DROP POLICY IF EXISTS "leitura_por_usuario"     ON pedidos;
DROP POLICY IF EXISTS "escrita_por_usuario"      ON pedidos;
DROP POLICY IF EXISTS "atualizacao_por_usuario"  ON pedidos;
DROP POLICY IF EXISTS "exclusao_por_usuario"     ON pedidos;
DROP POLICY IF EXISTS "Acesso público"           ON pedidos;
DROP POLICY IF EXISTS "leitura_publica"          ON pedidos;
DROP POLICY IF EXISTS "escrita_publica"          ON pedidos;
DROP POLICY IF EXISTS "atualizacao_publica"      ON pedidos;
DROP POLICY IF EXISTS "exclusao_publica"         ON pedidos;

ALTER TABLE pedidos ENABLE ROW LEVEL SECURITY;

CREATE POLICY "pedidos_select" ON pedidos FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "pedidos_insert" ON pedidos FOR INSERT WITH CHECK (auth.uid() = user_id);
CREATE POLICY "pedidos_update" ON pedidos FOR UPDATE USING (auth.uid() = user_id);
CREATE POLICY "pedidos_delete" ON pedidos FOR DELETE USING (auth.uid() = user_id);

-- 4. Recria todas as policies de devoluções
DROP POLICY IF EXISTS "dev_leitura"     ON devolucoes;
DROP POLICY IF EXISTS "dev_escrita"     ON devolucoes;
DROP POLICY IF EXISTS "dev_atualizacao" ON devolucoes;
DROP POLICY IF EXISTS "dev_exclusao"    ON devolucoes;

ALTER TABLE devolucoes ENABLE ROW LEVEL SECURITY;

CREATE POLICY "dev_select" ON devolucoes FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "dev_insert" ON devolucoes FOR INSERT WITH CHECK (auth.uid() = user_id);
CREATE POLICY "dev_update" ON devolucoes FOR UPDATE USING (auth.uid() = user_id);
CREATE POLICY "dev_delete" ON devolucoes FOR DELETE USING (auth.uid() = user_id);

-- 5. Garante que a constraint unique existe
ALTER TABLE pedidos DROP CONSTRAINT IF EXISTS pedidos_user_pedido;
ALTER TABLE pedidos ADD CONSTRAINT pedidos_user_pedido UNIQUE (user_id, id_pedido);

ALTER TABLE devolucoes DROP CONSTRAINT IF EXISTS devolucoes_user_pedido;
ALTER TABLE devolucoes ADD CONSTRAINT devolucoes_user_pedido UNIQUE (user_id, id_pedido);

-- 6. Garante colunas
ALTER TABLE pedidos ADD COLUMN IF NOT EXISTS status_interno text DEFAULT '';
ALTER TABLE pedidos ADD COLUMN IF NOT EXISTS nota_revisao  text DEFAULT '';
ALTER TABLE pedidos ADD COLUMN IF NOT EXISTS loja          text DEFAULT '';
ALTER TABLE pedidos ADD COLUMN IF NOT EXISTS preco         numeric(12,2) DEFAULT 0;
