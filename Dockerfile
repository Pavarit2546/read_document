# ใช้ Python 3.9-slim เป็น Base Image
FROM python:3.9-slim

WORKDIR /app

COPY requirement.txt .

RUN pip install --no-cache-dir --upgrade pip -r requirement.txt

COPY . .

EXPOSE 3000

CMD ["gunicorn", "--bind", "0.0.0.0:3000", "app:app"]