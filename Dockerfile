FROM python:3.10-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY App.py .

# Render / Railway 會注入 PORT
ENV PORT=10000
EXPOSE $PORT

CMD gunicorn --bind 0.0.0.0:$PORT --workers 2 --timeout 120 App:server