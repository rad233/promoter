export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  if (req.method !== 'POST') { res.status(405).json({ error: 'Method not allowed' }); return; }

  try {
    const { pdfBase64, prompt } = req.body;
    if (!pdfBase64 || !prompt) {
      res.status(400).json({ error: 'Missing pdfBase64 or prompt' });
      return;
    }

    const GEMINI_KEY = process.env.GEMINI_API_KEY;
    if (!GEMINI_KEY) {
      res.status(500).json({ error: 'API key not configured' });
      return;
    }

    const MODELS = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash'];

    async function callGemini(promptText) {
      const MAX_RETRIES = 3;

      for (const model of MODELS) {
        for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
          try {
            const response = await fetch(
              `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_KEY}`,
              {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  contents: [{
                    parts: [
                      { inline_data: { mime_type: 'application/pdf', data: pdfBase64 } },
                      { text: promptText }
                    ]
                  }],
                  generationConfig: {
                    temperature: 0.1,
                    maxOutputTokens: 65536,
                    response_mime_type: 'application/json'
                  }
                })
              }
            );
            const data = await response.json();
            // If overloaded (503) or rate limited (429), retry
            if (response.status === 503 || response.status === 429) {
              const waitMs = attempt * 5000; // 5s, 10s, 15s
              console.warn(`Model ${model} overloaded (attempt ${attempt}), retrying in ${waitMs}ms...`);
              await sleep(waitMs);
              continue;
            }
            if (!response.ok) throw new Error(data.error?.message || 'Gemini API error');
            return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
          } catch (e) {
            if (attempt === MAX_RETRIES) {
              console.warn(`Model ${model} failed after ${MAX_RETRIES} attempts, trying next model...`);
              break; // try next model
            }
            await sleep(attempt * 5000);
          }
        }
      }
      throw new Error('All Gemini models are currently unavailable. Please try again in a few minutes.');
    }

    function cleanAndParse(raw) {
      raw = (raw || '').replace(/```json|```/g, '').trim();
      const js = raw.indexOf('{');
      const je = raw.lastIndexOf('}');
      if (js === -1 || je === -1) throw new Error('No JSON found');
      return JSON.parse(raw.slice(js, je + 1));
    }

    function sleep(ms) {
      return new Promise(resolve => setTimeout(resolve, ms));
    }

    // ── Call 1: Structure (questions, options, dialogue, word bank) ───────────
    const prompt1 = `Parse this exam PDF. Return ONLY valid JSON. Extract questions and structure but leave article texts empty for now:
{
  "hasIntro": true,
  "introText": "full intro page text",
  "tasks": [
    {"task":1,"type":"listening","points":8,"question_count":8,"instructions":"instruction text","questions":["q1"],"options":[["A. opt","B. opt","C. opt","D. opt"]]},
    {"task":2,"type":"match","points":8,"question_count":8,"instructions":"instruction text","questions":["statement1"],"passages":{"A":"full paragraph A","B":"full paragraph B","C":"paragraph C","D":"paragraph D","E":"paragraph E","F":"paragraph F"}},
    {"task":3,"type":"reading","points":8,"question_count":8,"instructions":"instruction text","text":"","questions":["q1"],"options":[["A. opt","B. opt","C. opt","D. opt"]]},
    {"task":4,"type":"gapfill4","points":12,"question_count":12,"instructions":"instruction text","text":"","word_bank":{"A":"word","B":"word"}},
    {"task":5,"type":"gapfill5","points":12,"question_count":12,"instructions":"instruction text","text":"","choices":[["A. word","B. word","C. word","D. word"]]},
    {"task":6,"type":"dialogue","points":6,"question_count":6,"instructions":"instruction text","dialogue":["Speaker: line","Speaker: ...(1)"],"options":{"A":"sentence","B":"sentence"}},
    {"task":7,"type":"essay","points":16,"instructions":"instruction text","prompt":"essay question"}
  ]
}
RULES: Preserve Georgian text exactly. Extract ALL tasks. Include full match paragraphs. Leave text field empty for reading/gapfill. Detect task type from content.`;

    const raw1 = await callGemini(prompt1);
    const parsed1 = cleanAndParse(raw1);

    // ── Wait 3 seconds between calls to avoid rate limit ─────────────────────
    await sleep(3000);

    // ── Call 2: Long texts only (reading passage + gapfill articles) ─────────
    const prompt2 = `From this exam PDF extract ONLY the long article/passage texts. Return ONLY valid JSON:
{
  "texts": {
    "task3": "COMPLETE reading passage word for word - mandatory, do not leave empty",
    "task4": "COMPLETE article with gap markers ......(1) ......(2) etc at exact positions - mandatory",
    "task5": "COMPLETE article with gap markers ......(1) ......(2) etc - mandatory"
  }
}
Copy text EXACTLY. Preserve Georgian. Include gap markers at exact positions.`;

    const raw2 = await callGemini(prompt2);
    let parsed2;
    try { parsed2 = cleanAndParse(raw2); } catch(e) { parsed2 = { texts: {} }; }

    // ── Merge texts into main result ─────────────────────────────────────────
    if (parsed2.texts && parsed1.tasks) {
      for (let i = 0; i < parsed1.tasks.length; i++) {
        const t = parsed1.tasks[i];
        const key = 'task' + t.task;
        if (parsed2.texts[key]) t.text = parsed2.texts[key];
      }
    }

    res.status(200).json({ text: JSON.stringify(parsed1) });

  } catch (err) {
    console.error('Parse error:', err);
    res.status(500).json({ error: err.message || 'Internal server error' });
  }
}
