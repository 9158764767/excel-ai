"""
A simple chatbot-like application for spreadsheet analysis
================================================================

This module defines a FastAPI application that allows a user to upload
an Excel or CSV file and then interact with it via a chat-like
interface. Unlike a full language model, this bot uses simple rules
to interpret commands such as "summary", "clean", "describe",
"visualize", and "forecast". The results are returned either as
structured JSON, a cleaned spreadsheet, a ZIP archive of charts, or a
forecasted Excel file depending on the requested operation.

Usage
-----

Run this server locally with Uvicorn:

    uvicorn main:app --reload --host 0.0.0.0 --port 8000

Then open your browser to http://localhost:8000 to access the web
interface. Upload a dataset and chat with the bot using simple
commands (e.g. "summary", "clean", "describe", "visualize",
"forecast target=my_column").

Limitations
-----------

This bot does not perform natural-language understanding. It looks for
specific keywords in your chat messages and executes the corresponding
analysis. If a command isn't recognized it will respond with a help
message. To build a more sophisticated agent you could replace the
command parser with a language model via the OpenAI API, but that
approach is beyond the scope of this example.
"""

import io
import os
from typing import Optional

from fastapi import FastAPI, UploadFile, Form, Request, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
import matplotlib.pyplot as plt
import zipfile

# Use a non-interactive backend for matplotlib
import matplotlib
matplotlib.use("Agg")

app = FastAPI(title="Agentic Excel AI")

# Mount static and template directories. These directories are created
# below in this patch to hold the HTML and JS for the chat interface.
BASE_DIR = os.path.dirname(__file__)
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATES_DIR)

# Store the uploaded DataFrame in memory. In a production system you
# would likely use a database or session store.
_dataframe: Optional[pd.DataFrame] = None

# ---------- Helper functions ----------

def load_dataframe(file: UploadFile) -> pd.DataFrame:
    """Load an Excel or CSV file into a pandas DataFrame."""
    content = file.file.read()
    filename = file.filename.lower()
    try:
        if filename.endswith((".xlsx", ".xls")):
            df = pd.read_excel(io.BytesIO(content))
        elif filename.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(content))
        else:
            raise ValueError("Unsupported file format. Please upload .xls/.xlsx or .csv")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {exc}") from exc
    return df


def analyze_summary(df: pd.DataFrame) -> dict:
    """Return summary statistics for the DataFrame."""
    return {
        "shape": {"rows": int(df.shape[0]), "columns": int(df.shape[1])},
        "columns": df.columns.tolist(),
        "missing_values": int(df.isnull().sum().sum()),
        "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
    }


def analyze_clean(df: pd.DataFrame) -> StreamingResponse:
    """Fill missing values and return a cleaned Excel file."""
    cleaned = df.copy()
    for col in cleaned.columns:
        if cleaned[col].dtype.kind in "biufc":
            cleaned[col] = cleaned[col].fillna(cleaned[col].median())
        else:
            mode_series = cleaned[col].mode()
            cleaned[col] = cleaned[col].fillna(mode_series.iloc[0] if not mode_series.empty else "")
    buf = io.BytesIO()
    cleaned.to_excel(buf, index=False)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=cleaned_data.xlsx"},
    )


def analyze_describe(df: pd.DataFrame) -> dict:
    """Return descriptive statistics and correlation matrix."""
    describe = df.describe(include="all").fillna("")
    corr = df.corr(numeric_only=True).fillna("")
    return {
        "describe": describe.to_dict(),
        "correlation": corr.to_dict(),
    }


def analyze_visualize(df: pd.DataFrame) -> StreamingResponse:
    """Generate histograms for up to three numeric columns and return a ZIP."""
    numeric_cols = df.select_dtypes(include="number").columns
    if len(numeric_cols) == 0:
        raise HTTPException(status_code=400, detail="No numeric columns to visualize.")
    numeric_cols = numeric_cols[:3]
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w") as zf:
        for col in numeric_cols:
            fig, ax = plt.subplots()
            df[col].dropna().plot(kind="hist", ax=ax)
            ax.set_title(col)
            ax.set_xlabel(col)
            ax.set_ylabel("Frequency")
            img_buf = io.BytesIO()
            fig.savefig(img_buf, format="png")
            plt.close(fig)
            img_buf.seek(0)
            zf.writestr(f"{col}.png", img_buf.read())
    zip_buf.seek(0)
    return StreamingResponse(
        zip_buf,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=charts.zip"},
    )


def analyze_forecast(df: pd.DataFrame, target: str) -> StreamingResponse:
    """Generate a simple forecast by predicting the mean of a target column."""
    if target not in df.columns:
        raise HTTPException(status_code=400, detail=f"Column '{target}' not found in dataset.")
    if not pd.api.types.is_numeric_dtype(df[target]):
        raise HTTPException(status_code=400, detail=f"Column '{target}' must be numeric for forecasting.")
    mean_value = float(df[target].mean())
    forecast_df = pd.DataFrame({target: [mean_value]})
    buf = io.BytesIO()
    forecast_df.to_excel(buf, index=False)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=forecast_{target}.xlsx"},
    )


def process_command(command: str) -> dict | StreamingResponse:
    """Parse a chat message and dispatch to the appropriate analysis function."""
    global _dataframe
    if _dataframe is None:
        return {"message": "No data uploaded yet. Please upload a dataset first."}
    cmd = command.strip().lower()
    if cmd.startswith("summary"):
        return {"summary": analyze_summary(_dataframe)}
    elif cmd.startswith("clean"):
        return analyze_clean(_dataframe)
    elif cmd.startswith("describe"):
        return analyze_describe(_dataframe)
    elif cmd.startswith("visualize"):
        return analyze_visualize(_dataframe)
    elif cmd.startswith("forecast"):
        # Extract target from command: e.g. "forecast target=column_name"
        parts = cmd.split()
        target = None
        for part in parts[1:]:
            if part.startswith("target="):
                target = part[len("target="):]
                break
        if not target:
            return {"message": "Please specify the target column, e.g. 'forecast target=Sales'"}
        return analyze_forecast(_dataframe, target)
    else:
        return {"message": "Unrecognized command. Available commands: summary, clean, describe, visualize, forecast target=<column>"}

# ---------- Routes ----------

@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    """Serve the main page with upload and chat UI."""
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/upload")
async def upload_file(file: UploadFile) -> JSONResponse:
    """Handle dataset upload and store the DataFrame in memory."""
    global _dataframe
    df = load_dataframe(file)
    _dataframe = df
    return JSONResponse({"message": f"Uploaded '{file.filename}' successfully! Now you can send commands."})


@app.post("/chat")
async def chat(message: str = Form(...)):
    """Handle chat messages and return the analysis result."""
    result = process_command(message)
    # If the result is a StreamingResponse, return it directly
    if isinstance(result, StreamingResponse):
        return result
    # Otherwise return JSON
    return JSONResponse(result)
