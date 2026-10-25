# Can I Drive Here? A RAG chatbot for driving in Victoria

A free chatbot that answers international students' questions about driving in Victoria, Australia (overseas licence rules, road rules, signs, parking, fines). It answers **only** from a curated knowledge base and says "I don't know" otherwise.

WIL group project, "RAG Model Blueprint" team. Trello board: https://trello.com/b/UryKBIB6/rag-model-blueprint-%F0%9F%A4%96

| Member | Role |
|---|---|
| Abhinav Velath Baiju (AV) | Lead, planning, coding |
| Akash Tony (AT) | Chunking, AI, coding |
| Ajith Kadavil Anilkumar (AK) | Chunking, AI, coding |
| Devika Vijayakumar Nair (DV) | Preprocessing knowledge base, coding |
| Vaishnav Vinod Nair (VV) | Preprocessing knowledge base, coding |

## How it works
Word knowledge base -> validated chunks (JSON) -> embeddings (`BAAI/bge-small-en-v1.5`) in Chroma -> top 5 chunks -> a free LLM (Groq first, Google Gemini as backup) told to answer only from them.

## Results (our own run, 47 student-style questions)
- Right chunk in the top 1 / 3 / 5: **71% / 86% / 95%** (42 in-scope questions)
- Answers passing the automatic check: **38 of 47 (81%)**, including all licence-type and out-of-scope questions
- Limits: small team-written test set; some failures are strict wording checks; known gaps (phone fine amount, tram-lane 50 m exception). Details in `docs/EVALUATION_REPORT.md` and `docs/PROJECT_REPORT.md`.

## Repository layout
```
Victoria_Driving_Chatbot.ipynb   the chatbot (run this)
data/
  integrated_kb_chunks.json      624 validated chunks the notebook loads
  evaluation_questions.json      47 test questions with expected facts
knowledge-base/
  Integrated_KB_self_contained.docx   main knowledge base (88 sections)
  URL_KB_self_contained.docx          government web pages
  PDF_KB_self_contained.docx          handbook
scripts/extract_and_chunk.py     Word file -> validated chunk JSON
docs/                            evaluation report, rewrite changelog, project summary, presentation
```

## Run it
1. `pip install -r requirements.txt`
2. Get a free key from Groq (https://console.groq.com/keys) and/or Google AI Studio (https://aistudio.google.com/apikey).
3. Set it as an environment variable, **never in the notebook or the repo**:
   `export GROQ_API_KEY=...` (and optionally `export GEMINI_API_KEY=...`). If none is set, the notebook asks for a Groq key with a hidden prompt.
4. Open `Victoria_Driving_Chatbot.ipynb` from the repo root and run the cells top to bottom. Steps 1 to 4 need no key.

## Rebuild the chunks
```
python scripts/extract_and_chunk.py knowledge-base/Integrated_KB_self_contained.docx data/integrated_kb_chunks.json
python scripts/extract_and_chunk.py --test      # run the script's own tests
```
The script refuses to write output if a validation check fails.

## Source material
Content is drawn from Victorian government web pages and *Road to Solo Driving* (Department of Transport and Planning, April 2023). The handbook PDF is not included; get it from the official site. Rules change, so confirm with VicRoads / Transport Victoria before relying on an answer.

## Honest limits
Not tested with real students. The knowledge base needs re-checking as rules and fees change.
