FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY splitter.py config.json ./

RUN mkdir -p /app/input /app/output

ENTRYPOINT ["python3", "splitter.py"]
