# lightweight Python base
FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive

# Install system deps and Tesseract (and English language pack)
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      build-essential \
      tesseract-ocr \
      libtesseract-dev \
      libleptonica-dev \
      poppler-utils \
      tesseract-ocr-eng \
      ca-certificates \
      && rm -rf /var/lib/apt/lists/*

# Create workspace//
WORKDIR /workspace

# copy requirements later via devcontainer postCreateCommand, but include default
COPY requirements.txt /workspace/requirements.txt

# Install python deps
RUN pip install --no-cache-dir -r /workspace/requirements.txt

CMD [ "bash" ]
