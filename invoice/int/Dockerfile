FROM python:3.11-slim

# Supaya log python langsung muncul di docker logs, tanpa buffering
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# curl dibutuhkan untuk HEALTHCHECK
RUN apt-get update \
    && apt-get install -y --no-install-recommends curl \
    && rm -rf /var/lib/apt/lists/*

# Install dependencies dulu (biar layer cache-nya efisien saat rebuild)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy semua source code (app.py, Template_US.xlsx, Template_Int.xlsx, dll)
COPY . .

EXPOSE 8501

HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

ENTRYPOINT ["streamlit", "run", "app.py", \
    "--server.port=8501", \
    "--server.address=0.0.0.0"]