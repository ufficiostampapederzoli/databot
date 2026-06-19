const express  = require("express");
const path     = require("path");
const fs       = require("fs");
const XLSX     = require("xlsx");
const Anthropic = require("@anthropic-ai/sdk");

const app    = express();
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── Protezione con password (opzionale) ────────────────────────────────────────
function checkAuth(req, res, next) {
  const user = process.env.SITE_USERNAME;
  const pass = process.env.SITE_PASSWORD;
  if (!user || !pass) return next();
  const header = req.headers.authorization;
  if (header) {
    const decoded = Buffer.from(header.split(" ")[1] || "", "base64").toString();
    const [u, p] = decoded.split(":");
    if (u === user && p === pass) return next();
  }
  res.set("WWW-Authenticate", 'Basic realm="NINA"');
  return res.status(401).send("Accesso richiesto.");
}

app.use(checkAuth);
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

// ── Caricamento dati in memoria ────────────────────────────────────────────────
// I dati vengono letti UNA VOLTA sola all'avvio e tenuti in memoria.
// Questo è fondamentale per il Prompt Caching: il testo deve essere
// identico tra una richiesta e l'altra perché il cache scatti.
let DATA_CONTEXT = "";

function buildDataContext() {
  const dataDir = path.join(__dirname, "data");
  if (!fs.existsSync(dataDir)) return "(Nessun dato disponibile)";

  const files = fs.readdirSync(dataDir).filter(f => /\.xlsx?$/i.test(f));
  if (!files.length) return "(Nessun dato disponibile)";

  let ctx = "";
  for (const filename of files) {
    try {
      const wb = XLSX.readFile(path.join(dataDir, filename));
      ctx += `\n# ${filename}\n`;
      for (const sheetName of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
        if (!rows.length) continue;

        ctx += `\n## ${sheetName} (${rows.length} righe)\n`;

        // Formato CSV compatto: riduce i token del ~35% rispetto al formato chiave:valore
        const headers = Object.keys(rows[0]);
        ctx += headers.join(",") + "\n";

        const MAX = 500;
        rows.slice(0, MAX).forEach(row => {
          const line = headers.map(h => {
            const v = String(row[h] ?? "").replace(/,/g, ";");
            return v;
          }).join(",");
          ctx += line + "\n";
        });
        if (rows.length > MAX) ctx += `[Prime ${MAX} righe su ${rows.length}]\n`;
      }
    } catch (e) {
      ctx += `\n[Errore lettura ${filename}: ${e.message}]\n`;
    }
  }
  return ctx;
}

function loadData() {
  DATA_CONTEXT = buildDataContext();
  const dataDir = path.join(__dirname, "data");
  if (fs.existsSync(dataDir)) {
    const files = fs.readdirSync(dataDir).filter(f => /\.xlsx?$/i.test(f));
    console.log(files.length > 0
      ? `📊 Dati caricati: ${files.join(", ")}`
      : "⚠️  Nessun file Excel nella cartella /data");
  }
}

// ── System prompt ──────────────────────────────────────────────────────────────
const SYSTEM_INSTRUCTIONS = `Sei NINA (Network Informativo per i Numeri dell'Assistenza Sanitaria), un assistente professionale per l'analisi di dati sanitari strutturati.
Rispondi SEMPRE in italiano con linguaggio chiaro e professionale.
Basati ESCLUSIVAMENTE sui dati forniti. Se un dato non è presente, dichiaralo esplicitamente senza inventare.
Non rivelare mai i nomi dei file sorgente né la struttura interna dei dati.
Se ti viene chiesto della fonte, rispondi: "I dati provengono da fonti ufficiali aggiornate periodicamente."
Puoi calcolare percentuali, medie, variazioni anno su anno, trend e classifiche.
Quando esegui calcoli, mostra brevemente il ragionamento.
Cita sempre il periodo di riferimento quando disponibile.

REGOLE SUL DOWNLOAD E L'ESPORTAZIONE DEI DATI:
Non puoi generare, trasmettere o esportare file di alcun tipo.
Non suggerire mai di copiare tabelle dalla chat per creare file.
Se l'utente chiede di scaricare o esportare dati, rispondi: "Per ottenere i dati originali ti invito a contattare la fonte ufficiale di riferimento."`;

// ── Endpoint chat ─────────────────────────────────────────────────────────────
app.post("/api/chat", async (req, res) => {
  try {
    if (!process.env.ANTHROPIC_API_KEY) {
      return res.status(500).json({ error: "Chiave API non configurata." });
    }

    const { messages } = req.body;
    if (!messages || !Array.isArray(messages)) {
      return res.status(400).json({ error: "Formato richiesta non valido." });
    }

    // System prompt con cache_control:
    // Il blocco "istruzioni + dati" viene memorizzato da Anthropic.
    // Le richieste successive che trovano lo stesso testo in cache
    // costano il 90% in meno sull'input.
    const systemPrompt = [
      {
        type: "text",
        text: `${SYSTEM_INSTRUCTIONS}\n\n=== DATI DISPONIBILI ===\n${DATA_CONTEXT}`,
        cache_control: { type: "ephemeral" }
      }
    ];

    const response = await client.messages.create({
      model:      "claude-sonnet-4-6",
      max_tokens: 4000,
      system:     systemPrompt,
      messages,
    });

    // Log dell'utilizzo cache (utile per monitorare i risparmi)
    const usage = response.usage;
    if (usage) {
      const cached = usage.cache_read_input_tokens || 0;
      const total  = usage.input_tokens || 0;
      if (cached > 0) {
        console.log(`💾 Cache: ${cached}/${total} token da cache (${Math.round(cached/total*100)}% risparmiato)`);
      }
    }

    res.json({ content: response.content[0].text });
  } catch (err) {
    console.error("Errore API:", err.message);
    res.status(500).json({ error: "Errore interno. Riprova tra qualche istante." });
  }
});

// ── Avvio server ──────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ NINA avviata sulla porta ${PORT}`);
  loadData();
});
