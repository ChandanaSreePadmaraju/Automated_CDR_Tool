FROM python:3.11-slim

# System libraries required by lxml and python-docx
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        gcc \
        libxml2-dev \
        libxslt1-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies first so this layer is cached separately
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Pre-download the BAAI/bge-base-en-v1.5 model into the image so that
# first-run startup is instant.  Requires network access during 'docker build'.
# Remove this RUN step if building in an air-gapped environment and mount the
# HuggingFace cache as a volume instead.
ENV HF_HUB_DISABLE_SSL_VERIFICATION=1
RUN python -c "\
from sentence_transformers import SentenceTransformer; \
SentenceTransformer('BAAI/bge-base-en-v1.5')"

# Copy application source and configuration
COPY . .

EXPOSE 8501

ENV STREAMLIT_SERVER_ADDRESS=0.0.0.0
ENV STREAMLIT_SERVER_PORT=8501

CMD ["streamlit", "run", "app.py", \
     "--server.address=0.0.0.0", \
     "--server.port=8501", \
     "--server.headless=true"]
