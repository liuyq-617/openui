FROM ghcr.io/open-webui/open-webui:main

RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice fonts-noto-cjk && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

COPY doc_revision_engine.py /app/backend/open_webui/utils/doc_revision_engine.py
COPY backend/open_webui/tools/builtin.py /app/backend/open_webui/tools/builtin.py
COPY backend/open_webui/utils/tools.py /app/backend/open_webui/utils/tools.py
