-- Execute no Supabase > SQL Editor > New Query
-- Corrige os status antigos que foram salvos truncados

-- Ver quais status existem atualmente
SELECT status_pedido, COUNT(*) as total
FROM pedidos
WHERE user_id = '3ccdd460-3798-4d87-9571-33c16be7071f'
GROUP BY status_pedido
ORDER BY total DESC;
