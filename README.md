# 📦 Gerenciador de Pedidos

Sistema para unificar planilhas de pedidos do Shopee/UpSeller com persistência em nuvem.

## Stack

- **React + Vite** — interface
- **Supabase** — banco de dados PostgreSQL na nuvem
- **Vercel** — hospedagem com deploy automático via GitHub
- **SheetJS (xlsx)** — leitura e exportação de planilhas

---

## 1. Supabase — criar o banco de dados

1. Acesse [supabase.com](https://supabase.com) e crie uma conta
2. Clique em **New Project** e dê um nome
3. Vá em **SQL Editor → New Query**
4. Cole o conteúdo do arquivo `supabase_setup.sql` e clique em **Run**
5. Vá em **Project Settings → API** e copie:
   - **Project URL**
   - **anon / public key**

---

## 2. Configurar variáveis de ambiente

Crie um arquivo `.env` na raiz do projeto (não sobe ao GitHub):

```
VITE_SUPABASE_URL=https://SEU-PROJETO.supabase.co
VITE_SUPABASE_ANON_KEY=sua-anon-key
```

---

## 3. GitHub — subir o código

```bash
# Instale o Git: https://git-scm.com
git init
git add .
git commit -m "primeiro commit"

# Crie um repositório no github.com e conecte:
git remote add origin https://github.com/SEU-USUARIO/gerenciador-pedidos.git
git branch -M main
git push -u origin main
```

---

## 4. Vercel — publicar com deploy automático

1. Acesse [vercel.com](https://vercel.com) e crie conta (pode entrar com GitHub)
2. Clique em **Add New → Project**
3. Importe o repositório `gerenciador-pedidos` do GitHub
4. Em **Environment Variables** adicione:
   - `VITE_SUPABASE_URL` → sua URL
   - `VITE_SUPABASE_ANON_KEY` → sua chave
5. Clique em **Deploy**

A cada `git push`, o Vercel republica automaticamente.

---

## Rodar localmente

```bash
npm install
npm run dev
# Acesse http://localhost:5173
```
