FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /workspace
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
# Northflank zwykle wystawia 8080
ENV PORT=8080
EXPOSE 8080

CMD ["python", "raporty_bot.py"]
