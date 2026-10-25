import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const GEMINI_API_KEY = Deno.env.get("GEMINI_API_KEY");
const API_BASE = "https://generativelanguage.googleapis.com/v1beta";
const PREFERRED_MODEL = Deno.env.get("GEMINI_MODEL") || "gemini-3.8-flash";
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
// O JSON (com a imagem em base64) é montado uma vez só por variante, para poupar memória.
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

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  try {
    if (!GEMINI_API_KEY) throw new Error("GEMINI_API_KEY não configurada");
    const { mime_type, data, modo } = await req.json();
    if (!mime_type || !data) throw new Error("Arquivo não enviado");
    if (!/^(image\/(jpeg|png|webp|heic|heif)|application\/pdf)$/.test(mime_type)) throw new Error("Formato não suportado. Envie foto (JPG/PNG) ou PDF.");
    const variacoes = modo === "variacoes";

    const payload = {
      contents: [{ role: "user", parts: [{ inline_data: { mime_type, data } }, { text: variacoes ? PROMPT_VARIACOES : PROMPT }] }],
      generationConfig: { maxOutputTokens: 4096, temperature: 0, responseMimeType: "application/json" },
    };

    // Modelo sobrecarregado (503/429/500): tenta de novo e, se continuar, usa outro modelo flash
    const tentarModelo = async (m: string) => {
      let r = await chamar(m, payload);
      for (const espera of [1500]) {
        if (![429, 500, 503].includes(r.status)) break;
        await new Promise((res) => setTimeout(res, espera));
        r = await chamar(m, payload);
      }
      return r;
    };
    let model = cachedModel || PREFERRED_MODEL;
    let response: Response | null = null;
    // Tabela simples de variações: modelo "lite" é bem mais rápido; se falhar, segue o caminho normal
    if (variacoes) {
      try {
        const lite = await chamar("gemini-2.5-flash-lite", payload);
        if (lite.ok) { response = lite; model = "gemini-2.5-flash-lite"; }
      } catch (_) { /* segue para o modelo normal */ }
    }
    if (!response) response = await tentarModelo(model);
    if (!response.ok && [404, 429, 500, 503].includes(response.status)) {
      const original = response.status;
      const outros = (await listarModelosFlash()).filter((n) => n !== model).slice(0, 1);
      for (const alt of outros) {
        const r2 = await tentarModelo(alt);
        if (r2.ok) { response = r2; if (original === 404) cachedModel = alt; model = alt; break; }
      }
    }
    if (!response.ok) {
      const errText = await response.text();
      if ([429, 503].includes(response.status)) throw new Error("A leitura por IA está sobrecarregada agora. Aguarde um minuto e tente de novo.");
      throw new Error(`Gemini API error ${response.status} (modelo ${model}): ${errText}`);
    }

    const j = await response.json();
    const parts = j.candidates?.[0]?.content?.parts || [];
    let texto = parts.filter((p: any) => typeof p.text === "string" && !p.thought).map((p: any) => p.text).join("").trim();
    texto = texto.replace(/^```(?:json)?/i, "").replace(/```$/, "").trim();
    const ini = texto.indexOf("{"), fim = texto.lastIndexOf("}");
    if (ini < 0 || fim < 0) throw new Error(variacoes ? "Não consegui ler o print. Tente um print mais nítido." : "Não consegui ler a nota. Tente uma foto mais nítida ou o PDF.");
    const dados = JSON.parse(texto.slice(ini, fim + 1));

    return new Response(JSON.stringify({ dados }), { headers: { "Content-Type": "application/json", ...CORS } });
  } catch (err: any) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500, headers: { "Content-Type": "application/json", ...CORS } });
  }
});
