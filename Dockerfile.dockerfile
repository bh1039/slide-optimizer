FROM python:3.10-slim

# Install LibreOffice and clean up to keep the image small
RUN apt-get update && apt-get install -y \
    libreoffice-writer \
    libreoffice-impress \
    --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Start the web server using Gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "app:app"]