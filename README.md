# Excel AI Pro

A minimal FastAPI application that leverages OpenAI's Code Interpreter tool to
perform exploratory data analysis on uploaded spreadsheets. Users can choose
from a variety of tasks such as summarisation, cleaning, describing
relationships, visualising key columns, or forecasting. When applicable, a
cleaned spreadsheet is streamed back to the client.

## Features

- Upload `.xlsx`, `.xls`, or `.csv` files via a simple web interface.
- Task options: summarise, clean and fix types, describe correlations,
  visualise key columns, and forecast time-series targets.
- Returns either a JSON summary or streams back a cleaned Excel workbook.
- Optional simple API key authentication via the `X-App-Key` header.
- Dockerfile provided for containerised deployment.

## Quickstart (local)

```bash
# Clone the repository and navigate into the project directory
git clone <your-repo-url>
cd excel-ai

# Create a virtual environment (optional but recommended)
python -m venv .venv
source .venv/bin/activate  # On Windows use .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Set environment variables
export OPENAI_API_KEY=sk-...
export OPENAI_MODEL=gpt-4o
# Optional auth configuration
export REQUIRE_APP_KEY=1
export APP_KEY=change-this-key

# Run the server
./run_local.sh

# Then open http://localhost:8000 in your browser.
```

## Docker

```bash
# Build the image
docker build -t excel-ai .

# Run the container
docker run \
  -e OPENAI_API_KEY=sk-... \
  -e OPENAI_MODEL=gpt-4o \
  -e REQUIRE_APP_KEY=1 \
  -e APP_KEY=change-this-key \
  -p 8000:8000 \
  excel-ai

# Visit http://localhost:8000
```

## API Usage

The `/analyze` endpoint accepts multipart form data containing a file and a
`task` field. When `REQUIRE_APP_KEY=1`, include the `X-App-Key` header with
your secret value. The endpoint returns either an Excel workbook or a JSON
summary.

```bash
curl -X POST http://localhost:8000/analyze \
  -H "X-App-Key: change-this-key" \
  -F "task=clean" \
  -F "file=@/path/to/data.xlsx" \
  -o result.xlsx
```

## Notes

- The Code Interpreter tool may evolve; adjust the parsing logic in
  `app/main.py` if the structure of the `responses.create` output changes.
- Keep your OpenAI API key secure by never committing it to source control.
- You can integrate your own API calls inside `app/main.py` before or after
  calling OpenAI, for example to run additional analytics or persist results.
