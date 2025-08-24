# Running Excel-AI Locally

These steps will run the Excel-AI API on your machine without needing an external server.

## 1. Clone the repository

```bash
git clone https://github.com/9158764767/excel-ai.git
cd excel-ai
```

## 2. Create a virtual environment (recommended)

```bash
python -m venv .venv
source .venv/bin/activate    # On Windows: .venv\Scripts\activate
```

## 3. Install dependencies

Install the packages listed in `requirements.txt`, including FastAPI, Uvicorn, pandas, matplotlib, openpyxl, jinja2, and python-multipart:

```bash
pip install -r requirements.txt
```

## 4. Set environment variables (optional)

Excel-AI performs all analysis locally and does not call OpenAI, so an API key is not required. You may disable the app-key requirement:

```bash
export OPENAI_MODEL=gpt-4o
export REQUIRE_APP_KEY=0
# Optionally set APP_KEY to secure your API:
# export APP_KEY=my-secret-key
```

On Windows PowerShell, use `$env:OPENAI_MODEL`, etc., instead of `export`.

## 5. Run the server

Start the FastAPI app on port 8000:

```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

If you see an error about `python-multipart` being missing, install it:

```bash
pip install python-multipart
```

## 6. Use the app

1. Open your browser to `http://localhost:8000`.
2. Upload an Excel (`.xlsx/.xls`) or CSV file via the form.
3. Select a task: **summary**, **clean**, **describe**, **visualize**, or **forecast**.
4. View results on the page or download the returned file.

The server will run until you stop it with `Ctrl+C`.

---

These instructions are intended for development and local testing. For production deployments you’ll need a host that supports Python and FastAPI and appropriate environment variables.
