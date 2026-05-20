FROM python:3.11-slim

# Install system dependencies for Chromium
RUN apt-get update && apt-get install -y \
    chromium \
    chromium-driver \
    libnss3 \
    libnspr4 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libxkbcommon0 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxrandr2 \
    libgbm1 \
    libasound2 \
    libpango-1.0-0 \
    libcairo2 \
    fonts-liberation \
    wget \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright and its Chromium browser
RUN playwright install chromium
RUN playwright install-deps chromium

# Copy app files
COPY . .

# Expose Streamlit port
EXPOSE 8501

# Run the app
CMD streamlit run app.py \
    --server.port $PORT \
    --server.address 0.0.0.0 \
    --server.headless true
