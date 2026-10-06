DATA DETECTIVE AI - FINAL BUILD

1. Replace your project's server.js with this server.js.
2. Replace public/index.html with this public/index.html.
3. Keep your existing package.json/node_modules.
4. .env must contain OPENROUTER_API_KEY. You can optionally set OPENROUTER_MODEL to a model available in your OpenRouter account. Do not commit the key.
5. Run: npm run dev
6. Open: http://localhost:5000

Included behavior:
- Normal AI chat; no canned greeting/keyword replies.
- Tamil / Tanglish / English / mixed-language conversation.
- Semantic AI routing between normal chat and dataset analysis.
- AI-generated read-only DuckDB SQL from the real schema.
- SQL is returned in the API and shown in the frontend under Generated SQL.
- Verified DuckDB result is shown separately.
- AI-generated follow-up suggestions are shown as clickable chips.
- Browser voice input via SpeechRecognition when supported.
- CSV/XLSX/XLS/JSON upload.
- Follow-up context is preserved in the frontend.
- Result Excel export.
- No fake local fallback answer when the AI fails; the UI shows the real error.
