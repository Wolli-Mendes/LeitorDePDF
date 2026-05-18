FROM python:3.14-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY ["Leitor de PDF.py", "."]
COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

EXPOSE 5000
CMD ["python", "Leitor de PDF.py"]
