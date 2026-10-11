# FINOVA — Goal-based Personal Finance Planning Engine

A Flask web application starter built from the supplied 22-page `BƯỚC CODE.docx` specification. It uses a blue/lavender fintech visual system inspired by the supplied FINOVA reference website, with light/dark mode and scenic hero imagery.

## Features included

- Public landing page, guide, registration/login/logout, password reset request placeholder, responsive navigation, theme persistence.
- Customer profile and financial baseline entry; VND money fields formatted with comma separators and non-negative validation.
- Python financial summary: cash flow, expense/income ratio, DTI, savings capacity, assets, liabilities, estimated net worth.
- Five transparent scenario assumptions, heuristic suitability indicators, projection chart, risk alerts and general financial recommendations.
- What-if/stress test for income, expenses, unexpected expenses, debt repayments, and savings adjustments.
- Up to three goals, required monthly savings, feasibility indicator, goal ranking, hypothetical future-income planner, comparison chart.
- Excel `.xlsx` export for summary, scenarios, and goals.
- Four LLM task routes: parse/validate inputs, rank goals, explain stress-test trade-offs, and RAG-style support chatbot using FAQ knowledge base.
- Evaluation stars for Accuracy, Groundedness, Usability, Processing Time; average customer rating and criteria summary; support tickets and searchable FAQs.
- Server-side Google AI Studio key; passwords hashed; database queries parameterized.

## Deploy to Render Free

1. Create a GitHub repository and upload the contents of this folder.
2. In Render, choose **New + → Blueprint** and select the repository containing `render.yaml`, or create a **Web Service** with runtime **Python**, root directory set to the folder containing `app.py`, build command `pip install -r requirements.txt`, and start command `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 90`.
3. Select the **Free** instance plan.
4. Add `GOOGLE_AI_STUDIO_API_KEY` in Render's Environment settings if you want live Gemini responses. Leave it blank to use the built-in local fallback. Never place the key in HTML/JavaScript or commit it to Git.
5. Set `FLASK_SECRET_KEY` to a strong random value if not using the Blueprint-generated value.
6. Deploy. Render assigns the public URL after the build succeeds.

## Important free-tier limitations and production gaps

- Render Free web services may sleep after inactivity and have limited CPU/RAM/build/runtime resources. Free plans and provider policies can change; do not assume zero charges if you later upgrade or use paid external services. Gemini API quotas/pricing are separate from Render.
- The default SQLite database is stored at `/tmp/finova.db` and is **ephemeral** on Render Free. Data may be lost on restart/redeploy. For durable multi-user production use, connect a database with a suitable free tier, verify its current limits, and configure backups. Do not promise long-term persistence with this default.
- Password reset currently returns a safe placeholder because email delivery is not configured. A verified email service and one-time expiring reset tokens are required for a production reset workflow. Password change is intentionally not exposed as an insecure demo endpoint.
- Support attachments are shown in the UI but are not uploaded/stored in this starter.
- Accuracy, Groundedness, and Processing Time are not objectively benchmarked by user stars. Implement a test suite and server-side latency instrumentation before using these KPIs operationally.
- Scenario suitability scores are rule-based heuristics, not calibrated probabilities. Return assumptions are hypothetical, and all results are educational estimates.
- Before collecting real financial data, add CSRF protection, rate limiting, security headers, a privacy policy, database encryption/access controls, durable storage, account deletion, and security review.
- Scenic photos load from Unsplash URLs; external image availability and terms apply. Chart.js and Google Fonts are loaded from CDNs.

## Local run

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000`.

## API key safety

Set `GOOGLE_AI_STUDIO_API_KEY` as a server environment variable. The browser never receives the key. If the key is blank or the provider call fails, FINOVA returns clearly labeled local fallback guidance. Never commit `.env` or real keys.
