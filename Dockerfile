FROM python:3.13-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY server.py .

RUN mkdir -p /data && chown -R nobody:nogroup /app /data

ENV DOC_WORKSPACE_DIR=/data

EXPOSE 8000

USER nobody

CMD ["python", "server.py"]
