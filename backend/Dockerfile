FROM python:3.11-slim

# Install LibreOffice (untuk convert PDF) dan font yang dibutuhkan
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-roboto \
    fonts-montserrat \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy dependency dan install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy kode dan template
COPY app.py .
COPY template.xlsx .

# Expose port 8000
EXPOSE 8000

# Jalankan server
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]