# Jigyasa AI

Jigyasa AI is a student study companion for asking questions about study materials and practising with short quizzes. The app supports text extraction from PDF, PPTX, DOCX, and TXT files.

## Run locally

On Windows, run `.\run.ps1` from PowerShell and open <http://127.0.0.1:8123>. Keep the PowerShell window open while using the app.

The local build stores uploaded documents under the ignored `data/` folder. Quiz attempts, chat history, and progress are saved in browser storage on that device. The local server binds to `127.0.0.1`; do not expose this local, unauthenticated mode to the public internet.

## Hosted account and private storage setup

The hosted configuration requires Supabase Auth, a private Supabase Storage bucket, and the documents table. Production startup fails unless its required server-side configuration is present.

1. Create a Supabase project for Jigyasa AI.
2. In the Supabase SQL Editor, run [`supabase/migrations/202610110001_secure_student_data.sql`](./supabase/migrations/202610110001_secure_student_data.sql).
3. In Supabase Auth settings, enable email/password sign-in and require email confirmation for the public demo. Set the Site URL to the deployed Render origin so confirmation links return to Jigyasa.
4. In Render, create a Web Service from this repository and use the included `render.yaml` blueprint. Set these environment variables in the Render dashboard:
   - `APP_ORIGIN`: the exact deployed HTTPS origin, without a trailing slash.
   - `SUPABASE_URL`: the project URL.
   - `SUPABASE_ANON_KEY`: the project's publishable/anon key.
   - `SUPABASE_SERVICE_ROLE_KEY`: the server-only service role key. Never add this value to frontend code, GitHub, or chat.
   - `GEMINI_API_KEY`: the Jigyasa-only Gemini key. Keep it in Render's environment settings, never in GitHub or browser code.
5. Set Supabase Auth's Site URL to the Render URL after Render creates it, then redeploy if needed.

Hosted accounts use HttpOnly, HTTPS-only cookies. The browser cannot read the Supabase session tokens. The backend verifies each session with Supabase Auth and filters every document query by the verified account ID. The database tables have row-level security enabled and no direct browser-access policies; private files are only accessed by the backend. Never run this hosted service without the Supabase SQL setup and all required server-side secrets.

Each hosted account is limited to 25 documents, 100 MB of private study files, and 20 Gemini questions per UTC day. The question counter is updated atomically in Supabase so restarting the web service does not reset it. `JIGYASA_DAILY_AI_QUESTION_LIMIT` can be set to a value from 1 to 200 in Render.

If Gemini returns a temporary high-demand (HTTP 503) response, Jigyasa retries once after a short pause. If Gemini still fails, the hosted question counter is refunded. Apply new SQL migrations in `supabase/migrations/` to the Supabase project before deploying backend changes.

## Gemini usage

Gemini usage is governed by the current limits and billing status of the Google AI Studio project and selected model. Requests include only the selected materials' relevant extracted text. Larger documents or selected context can use more tokens. Check the active limits in [Google AI Studio](https://aistudio.google.com/rate-limit?timeRange=last-28-days); do not assume a fixed unlimited question count.

## Current scope and limits

- Supports PDF, PPTX, DOCX, and TXT files up to 25 MB each, within the hosted account's 100 MB storage allowance.
- Video transcription, legacy `.ppt`, scanned-PDF OCR, AI-generated quizzes, and cloud-synced quiz history are not implemented.
- Quiz attempts and learning history are scoped to an account in that browser, but are not synced across devices.
- Gemini answers are generated from selected study materials; the external Gemini service receives the question and selected text context.
- The Render free service may spin down when idle; its first request after inactivity can take longer.

## Validate backend tests

```powershell
.\.venv312\Scripts\python.exe -m pip install -r requirements-test.txt
.\.venv312\Scripts\python.exe -m unittest backend.test_main
```

## Existing Python learning example

The Iris/Logistic Regression example remains separate from the website:

```powershell
.\.venv312\Scripts\python.exe main.py
```
