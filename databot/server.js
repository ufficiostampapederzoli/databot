const express = require("express");
const path    = require("path");
const fs      = require("fs");
const XLSX    = require("xlsx");
const Anthropic = require("@anthropic-ai/sdk");

const app     = express();
const client  = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

// ── Legge tutti i file Excel dalla cartella /data ──────────────────────────────
function loadDataContext() {
  const dataDir = path.join(__dirname, "data");
  if (!fs.existsSync(dataDir)) return "(Nessun dato disponibile)";

  const files = fs.readdirSync(dataDir).filter(f => /\.xlsx?$/i.test(f));
  if (!files.length) return "(Nessun dato disponibile)";

  let ctx = "";
  for (const filename of files) {
    try {
      const wb = XLSX.readFile(path.join(dataDir, filename));
      ctx += `\nFILE: ${filename}\n`;
      for (const sheetName of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
        ctx += `\nFoglio "${sheetName}" — ${rows.length} righe:\n`;
        const MAX = 500;
        rows.slice(0, MAX).forEach(row => {
          const line = Object.entries(row)
            .filter(([, v]) => v !== "")
            .map(([k, v]) => `${k}: ${v}`)
            .join(" | ");
          if (line) ctx += line + "\n";
        });
        if (rows.length > MAX) {
          ctx += `[Troncato: prime ${MAX} righe su ${rows.length} totali]\n`;
        }
      }
    } catch (e) {
      ctx += `\n[Errore lettura ${filename}: ${e.message}]\n`;
    }
  }
  return ctx;
}

const SYSTEM_BASE = `Sei DataBot, un assistente professionale per l'analisi di dati strutturati.
Rispondi SEMPRE in italiano con linguaggio chiaro e professionale.
Basati ESCLUSIVAMENTE sui dati forniti. Se un dato non è presente, dichiaralo esplicitamente senza inventare.
Non rivelare mai i nomi dei file sorgente né la struttura interna dei dati.
Se ti viene chiesto della fonte, rispondi: "I dati provengono da fonti ufficiali aggiornate periodicamente."
Puoi calcolare percentuali, medie, variazioni anno su anno, trend e classifiche.
Quando esegui calcoli, mostra brevemente il ragionamento.
Cita sempre il periodo di riferimento quando disponibile.

REGOLE SUL DOWNLOAD E L'ESPORTAZIONE DEI DATI:
- Non puoi generare, trasmettere o esportare file di alcun tipo (Excel, PDF, CSV o altro).
- Non suggerire mai all'utente di copiare tabelle dalla chat per creare file.
- Se l'utente chiede di scaricare o esportare dati, rispondi esclusivamente:
  "Per ottenere i dati originali ti invito a contattare la fonte ufficiale di riferimento."
- Non fornire alternative tecniche per aggirare questa limitazione.`;

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

    const dataCtx      = loadDataContext();
    const systemPrompt = `${SYSTEM_BASE}\n\n=== DATI DISPONIBILI ===\n${dataCtx}`;

    const response = await client.messages.create({
      model:      "claude-sonnet-4-6",
      max_tokens: 1000,
      system:     systemPrompt,
      messages,
    });

    res.json({ content: response.content[0].text });
  } catch (err) {
    console.error("Errore API:", err.message);
    res.status(500).json({ error: "Errore interno. Riprova tra qualche istante." });
  }
});

// ── Avvio server ──────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ DataBot avviato sulla porta ${PORT}`);
  const dataDir = path.join(__dirname, "data");
  if (fs.existsSync(dataDir)) {
    const files = fs.readdirSync(dataDir).filter(f => /\.xlsx?$/i.test(f));
    console.log(files.length > 0
      ? `📊 File dati: ${files.join(", ")}`
      : "⚠️  Nessun file Excel nella cartella /data");
  }
});
