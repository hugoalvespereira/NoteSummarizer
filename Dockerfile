FROM node:22-bookworm-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOST=0.0.0.0 \
    PORT=8000 \
    DATA_DIR=/data \
    WORK_DIR=/tmp/powerpoint-notes-summarizer \
    HOME=/data \
    CODEX_HOME=/data/.codex \
    CODEX_BIN=/usr/local/bin/codex

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
      ca-certificates \
      git \
      libreoffice \
      python3 \
      python-is-python3 \
    && npm install -g @openai/codex \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

RUN mkdir -p /data /tmp/powerpoint-notes-summarizer \
    && chmod 700 /data

EXPOSE 8000
VOLUME ["/data"]

CMD ["python3", "app.py"]
