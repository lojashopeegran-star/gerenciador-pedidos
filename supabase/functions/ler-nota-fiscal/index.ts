import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const GEMINI_API_KEY = Deno.env.get("GEMINI_API_KEY");
const ANTHROPIC_API_KEY = Deno.env.get("ANTHROPIC_API_KEY"); // opcional: segunda IA (Claude) quando o Gemini está sobrecarregado
const CLAUDE_MODEL = Deno.env.get("ANTHROPIC_MODEL") || "claude-haiku-5-5";
const API_BASE = "https://generativelanguage.googleapis.com/v1beta";
const PREFERRED_MODEL = Deno.env.get("GEMINI_MODEL") || "gemini-3.8-flash";
const FAST_MODEL = Deno.env.get("GEMINI_FAST_MODEL") || "gemini-2.5-flash-lite";
let cachedModel: string | null = null;

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

const PROMPT = `Leia esta nota fiscal/cupom/orçamento brasileiro (foto ou PDF) e responda SOMENTE com JSON:
{"fornecedor":"nome do EMITENTE (quem vendeu)","cnpj":"só números ou vazio","numero_nota":"","data_emissao":"YYYY-MM-DD ou vazio","valor_total":0,"itens":[{"descricao":"","quantidade":0,"unidade":"UN","valor_unitario":0,"valor_total":0}]}
Regras: números com ponto decimal e sem milhar (1234.56); um item por linha de produto; não invente dados (use "" ou 0); valor_unitario = preço de UMA unidade (se faltar, total ÷ quantidade); frete, impostos e descontos não são itens; unidade como na nota (UN, CX, PCT, KG, L, M, RL).`;

// Print de lista de variações (ex.: tela da Shopee): 1ª coluna = nome da variação, ÚLTIMA coluna = preço
const PROMPT_VARIACOES = `Esta imagem é um print de uma tabela de variações de um produto (e-commerce). Responda SOMENTE com JSON:
{"variacoes":[{"nome":"","preco":0}]}
Regras: uma entrada por linha da tabela, na mesma ordem em que aparecem; "nome" = texto da PRIMEIRA coluna (esquerda), exatamente como está escrito; "preco" = valor da ÚLTIMA coluna (direita) como número com ponto decimal (12.99), ignorando as colunas do meio; ignore cabeçalhos, título/foto do produto e linhas sem nome; não invente linhas; se o preço da última coluna não estiver legível use 0.`;

async function callGemini(model: string, body: string) {
  return await fetch(`${API_BASE}/models/${model}:generateContent?key=${GEMINI_API_KEY}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body,
    signal: AbortSignal.timeout(40000), // não deixa uma chamada lenta travar a função inteira
  });
}

async function listarModelosFlash(): Promise<string[]> {
  const r = await fetch(`${API_BASE}/models?pageSize=200&key=${GEMINI_API_KEY}`);
  if (!r.ok) return [];
  const j = await r.json();
  const nomes: string[] = (j.models || [])
    .filter((m: any) => (m.supportedGenerationMethods || []).includes("generateContent"))
    .map((m: any) => String(m.name).replace("models/", ""))
    .filter((n: string) => /^gemini-[\d.]+-flash$/.test(n));
  const versao = (n: string) => parseFloat(n.match(/^gemini-([\d.]+)-flash$/)![1]);
  nomes.sort((a, b) => versao(b) - versao(a));
  return nomes;
}

// Tenta com "pensamento" desligado (bem mais rápido); se o modelo não aceitar, repete sem a opção.
async function chamar(model: string, base: any) {
  const corpo = (gc: any) => JSON.stringify({ ...base, generationConfig: { ...base.generationConfig, ...gc } });
  let r: Response;
  try { r = await callGemini(model, corpo({ thinkingConfig: { thinkingBudget: 0 } })); }
  catch (_) { return new Response("timeout", { status: 503 }); }
  if (r.status === 400) {
    const t = await r.clone().text();
    if (/thinking/i.test(t)) {
      try {
        r = await callGemini(model, corpo({ thinkingConfig: { thinkingLevel: "minimal" } }));
        if (r.status === 400) r = await callGemini(model, JSON.stringify(base));
      } catch (_) { return new Response("timeout", { status: 503 }); }
    }
  }
  return r;
}

// Extrai o JSON de um texto; devolve null se vier algo ilegível
function jsonDeTexto(texto: string): any | null {
  texto = texto.replace(/^```(?:json)?/i, "").replace(/```$/, "").trim();
  const ini = texto.indexOf("{"), fim = texto.lastIndexOf("}");
  if (ini < 0 || fim < 0) return null;
  try { return JSON.parse(texto.slice(ini, fim + 1)); } catch (_) { return null; }
}

function extrairJson(j: any): any | null {
  const parts = j?.candidates?.[0]?.content?.parts || [];
  const texto = parts.filter((p: any) => typeof p.text === "string" && !p.thought).map((p: any) => p.text).join("").trim();
  return jsonDeTexto(texto);
}

// Segunda IA: Claude (só roda se o secret ANTHROPIC_API_KEY existir)
async function lerComClaude(mime_type: string, data: string, prompt: string): Promise<any | null> {
  if (!ANTHROPIC_API_KEY) return null;
  const bloco = mime_type === "application/pdf"
    ? { type: "document", source: { type: "base64", media_type: mime_type, data } }
    : { type: "image", source: { type: "base64", media_type: mime_type, data } };
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "content-type": "application/json", "x-api-key": ANTHROPIC_API_KEY, "anthropic-version": "2023-06-01" },
    body: JSON.stringify({ model: CLAUDE_MODEL, max_tokens: 4096, temperature: 0, messages: [{ role: "user", content: [bloco, { type: "text", text: prompt }] }] }),
    signal: AbortSignal.timeout(45000),
  });
  if (!r.ok) throw new Error(`Claude API error ${r.status}: ${(await r.text()).slice(0, 300)}`);
  const j = await r.json();
  const texto = (j.content || []).filter((c: any) => c.type === "text").map((c: any) => c.text).join("").trim();
  return jsonDeTexto(texto);
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  try {
    if (!GEMINI_API_KEY && !ANTHROPIC_API_KEY) throw new Error("Nenhuma chave de IA configurada");
    const { mime_type, data, modo } = await req.json();
    if (!mime_type || !data) throw new Error("Arquivo não enviado");
    if (!/^(image\/(jpeg|png|webp|heic|heif)|application\/pdf)$/.test(mime_type)) throw new Error("Formato não suportado. Envie foto (JPG/PNG) ou PDF.");
    const variacoes = modo === "variacoes";
    const msgIlegivel = variacoes ? "Não consegui ler o print. Tente um print mais nítido." : "Não consegui ler a nota. Tente uma foto mais nítida ou o PDF.";
    const ok = (dados: any) => new Response(JSON.stringify({ dados }), { headers: { "Content-Type": "application/json", ...CORS } });
    const vazio = (d: any) => !d || (variacoes ? !(d.variacoes || []).length : !(d.itens || []).length);

    const payload = {
      contents: [{ role: "user", parts: [{ inline_data: { mime_type, data } }, { text: variacoes ? PROMPT_VARIACOES : PROMPT }] }],
      generationConfig: { maxOutputTokens: 4096, temperature: 0, responseMimeType: "application/json" },
    };

    let erroGemini = "";
    if (GEMINI_API_KEY) {
      // 1) modelo rápido (lite); se falhar ou vier ilegível, 2) modelo principal; 3) outro modelo flash
      try {
        const rapido = await chamar(FAST_MODEL, payload);
        if (rapido.ok) {
          const d = extrairJson(await rapido.json());
          if (!vazio(d)) return ok(d);
        }
      } catch (_) { /* segue para o modelo principal */ }

      const tentarModelo = async (m: string) => {
        let r = await chamar(m, payload);
        if ([429, 500, 503].includes(r.status)) {
          await new Promise((res) => setTimeout(res, 1500));
          r = await chamar(m, payload);
        }
        return r;
      };
      let model = cachedModel || PREFERRED_MODEL;
      let response = await tentarModelo(model);
      if (!response.ok && [404, 429, 500, 503].includes(response.status)) {
        const original = response.status;
        const outros = (await listarModelosFlash()).filter((n) => n !== model && n !== FAST_MODEL).slice(0, 1);
        for (const alt of outros) {
          const r2 = await tentarModelo(alt);
          if (r2.ok) { response = r2; if (original === 404) cachedModel = alt; model = alt; break; }
        }
      }
      if (response.ok) {
        const dados = extrairJson(await response.json());
        if (dados) return ok(dados);
        erroGemini = msgIlegivel;
      } else {
        const errText = await response.text();
        erroGemini = [429, 503].includes(response.status)
          ? "A leitura por IA está sobrecarregada agora. Aguarde um minuto e tente de novo."
          : `Gemini API error ${response.status} (modelo ${model}): ${errText}`;
      }
    }

    // 4) Gemini falhou/sobrecarregado: tenta o Claude (se a chave estiver configurada)
    if (ANTHROPIC_API_KEY) {
      try {
        const d = await lerComClaude(mime_type, data, variacoes ? PROMPT_VARIACOES : PROMPT);
        if (!vazio(d)) return ok(d);
        erroGemini = erroGemini || msgIlegivel;
      } catch (e: any) {
        erroGemini = (erroGemini ? erroGemini + " | " : "") + e.message;
      }
    }
    throw new Error(erroGemini || msgIlegivel);
  } catch (err: any) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { "Content-Type": "application/json", ...CORS } });
  }
});
