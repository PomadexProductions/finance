# Setup Guide — Pomadex Finance v10

## 1. GitHub Pages

```bash
# Clone ou cria o repo
git clone https://github.com/PomadexProductions/finance
cd finance

# Copia os ficheiros
cp index.html .
cp apps-script.js .
cp apple-touch-icon.png .  # já está no repo

# Commit e push
git add .
git commit -m "Finance v10 — sistema limpo com Sheets 2.0"
git push
```

Acesso: `https://pomadexproductions.github.io/finance/`

---

## 2. Vercel (alternativa mais rápida)

1. Vai a vercel.com → New Project
2. Import from GitHub → PomadexProductions/finance
3. Framework: Other (static)
4. Deploy!

URL: `https://pomadex-finance.vercel.app`

---

## 3. Google Apps Script

1. Vai a script.google.com
2. Novo projecto → nome: "Pomadex Finance API v2"
3. Cola o conteúdo de `apps-script.js`
4. Guarda (Ctrl+S)
5. Deploy → New deployment → Web App
   - Execute as: Me
   - Who has access: Anyone
6. Autoriza → Copia o URL `/exec`
7. No app → Settings → Cloud Sync → cola o URL
8. O sistema cria as abas automaticamente

---

## 4. Sheets — criar abas manualmente

No Sheets `19YutemqC9e3AQz3PR_Mfwn-c7GxgQ1PYd4vAwDRgWmE`:

Cria estas abas (exactamente com estes nomes):
- `Pomadex_Productions`
- `SOTS_Recordings`
- `Clube_do_Snare`
- `PEMAX`
- `Pedro_Maximo`
- `Categorias`
- `Dashboard`

**OU** — liga o app e faz uma primeira chamada ao Apps Script com action=setup
que cria tudo automaticamente.

---

## PIN default
`1234` — muda em Settings → Mudar PIN
