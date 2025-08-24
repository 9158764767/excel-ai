import io
import os
from fastapi import FastAPI, UploadFile, Form, Request, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import pandas as pd
import matplotlib.pyplot as plt
import zipfile

"""
This module defines a FastAPI application for performing Excel and CSV data analysis without
relying on external APIs or OpenAI's Code Interpreter. All operations are executed locally
using pandas and matplotlib. Supported tasks include summary statistics, cleaning missing
values, descriptive statistics, basic visualizations, and a trivial forecasting method.
"""

# We keep the OPENAI_MODEL variable for backwards compatibility with environments that may
# still reference it, but it is unused in this implementation.
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o")

def _get_app_key() -> str:
    """
    Retrieve the application key from the environment. This is used to enforce optional
    header-based authentication when the REQUIRE_APP_KEY environment variable is set to "1".

    Returns:
        The application key as a string, or an empty string if not set.
    """
    return os.environ.get("APP_KEY", "")

app = FastAPI(title="Excel AI Pro")

# Prepare and mount a static directory. Although it is currently unused, having the mount in
# place simplifies serving static assets in the future if needed.
static_dir = os.path.join(os.path.dirname(__file__), "static")
app.mount("/static", StaticFiles(directory=static_dir), name="static")

# Set up Jinja2 templates for the single-page HTML frontend.
templates = Jinja2Templates(directory=os.path.join(os.path.dirname(__file__), "templates"))

# Configure matplotlib to use a non-interactive backend. This is important in environments
# without display servers and prevents matplotlib from trying to open GUI windows.
import matplotlib  # imported here so it coexists with matplotlib.pyplot import above
matplotlib.use("Agg")

@app.get("/", response_class=HTMLResponse)
async def home(request: Request) -> HTMLResponse:
    """
    Render the homepage, which contains a form for uploading Excel/CSV files and selecting
    analysis tasks.
    """
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/health")
async def health_check() -> JSONResponse:
    """Simple health-check endpoint."""
    return JSONResponse({"status": "ok"})

@app.post("/analyze")
async def analyze_excel(
    request: Request,
    task: str = Form("summary"),
    file: UploadFile | None = None,
    target: str | None = Form(None),
) -> JSONResponse | StreamingResponse:
    """
    Perform the requested analysis task on an uploaded Excel or CSV file. All tasks are
    executed locally using pandas and matplotlib; no external services are called.

    Supported tasks:
      - summary: Return dataset shape, column names, missing value count and data types.
      - clean: Fill missing numeric values with the median and non-numeric with the mode,
               then return a cleaned Excel file.
      - describe: Return pandas describe() output and a correlation matrix.
      - visualize: Generate histograms for up to three numeric columns and return a ZIP
                   archive containing PNG images.
      - forecast: Simple forecasting for a numeric target column by predicting its mean.
    """
    if file is None:
        raise HTTPException(status_code=400, detail="No file provided")

    # Optional API key enforcement
    require_key = os.environ.get("REQUIRE_APP_KEY")
    if require_key == "1":
        header_key = request.headers.get("X-App-Key")
        if not header_key or header_key != _get_app_key():
            raise HTTPException(status_code=401, detail="Unauthorized")

    # Read the uploaded file into a pandas DataFrame
    file_bytes = await file.read()
    filename = file.filename.lower()
    try:
        if filename.endswith((".xlsx", ".xls")):
            df = pd.read_excel(io.BytesIO(file_bytes))
        elif filename.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(file_bytes))
        else:
            raise HTTPException(status_code=400, detail="Unsupported file type. Please upload an Excel or CSV file.")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {exc}")

    # Normalize task name
    task = task.lower().strip()

    if task == "summary":
        summary = {
            "shape": {"rows": int(df.shape[0]), "columns": int(df.shape[1])},
            "columns": df.columns.tolist(),
            "missing_values": int(df.isnull().sum().sum()),
            "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
        }
        return JSONResponse({"summary": summary})

    elif task == "clean":
        cleaned_df = df.copy()
        for col in cleaned_df.columns:
            if cleaned_df[col].dtype.kind in "biufc":
                median = cleaned_df[col].median()
                cleaned_df[col] = cleaned_df[col].fillna(median)
            else:
                mode_series = cleaned_df[col].mode()
                mode_value = mode_series.iloc[0] if not mode_series.empty else ""
                cleaned_df[col] = cleaned_df[col].fillna(mode_value)
        output = io.BytesIO()
        cleaned_df.to_excel(output, index=False)
        output.seek(0)
        headers = {"x-summary": f"Cleaned {file.filename}"}
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    elif task == "describe":
        describe_dict = df.describe(include="all").fillna("").to_dict()
        corr = df.corr(numeric_only=True).fillna("").to_dict()
        return JSONResponse({"describe": describe_dict, "correlation": corr})

    elif task == "visualize":
        numeric_cols = df.select_dtypes(include="number").columns
        if len(numeric_cols) == 0:
            raise HTTPException(status_code=400, detail="No numeric columns available for visualization.")
        numeric_cols = numeric_cols[:3]
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zip_file:
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
                zip_file.writestr(f"{col}.png", img_buf.read())
        zip_buf.seek(0)
        headers = {"Content-Disposition": f'attachment; filename="{file.filename.rsplit(".", 1)[0]}_charts.zip"'}
        return StreamingResponse(zip_buf, media_type="application/zip", headers=headers)

    elif task == "forecast":
        if target is None:
            raise HTTPException(status_code=400, detail="Please specify a target column for forecasting via the 'target' form field.")
        if target not in df.columns:
            raise HTTPException(status_code=400, detail=f"Target column '{target}' not found in the dataset.")
        if not pd.api.types.is_numeric_dtype(df[target]):
            raise HTTPException(status_code=400, detail=f"Target column '{target}' must be numeric for forecasting.")
        mean_value = float(df[target].mean())
        forecast_df = pd.DataFrame({target: [mean_value]})
        out = io.BytesIO()
        forecast_df.to_excel(out, index=False)
        out.seek(0)
        headers = {"x-summary": f"Forecasted mean for column '{target}'"}
        return StreamingResponse(
            out,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    else:
        raise HTTPException(status_code=400, detail=f"Unknown task '{task}'. Supported tasks are summary, clean, describe, visualize, forecast.")
