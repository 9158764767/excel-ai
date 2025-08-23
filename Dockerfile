FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
# Copy the application code into the image
COPY app ./app
# Default environment variables (can be overridden at runtime)
ENV HOST=0.0.0.0 \
    PORT=8000
EXPOSE 8000
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
