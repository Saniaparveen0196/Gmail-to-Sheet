# Use Python 3.9+ base image
FROM python:3.9-slim

# Set working directory
WORKDIR /app

# Create non-root user for security
RUN useradd -m -u 1000 appuser && chown -R appuser:appuser /app

# Install system dependencies if needed
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code
COPY config.py .
COPY src/ ./src/

# Create credentials and logs directories (will be mounted as volumes)
RUN mkdir -p credentials logs && chown -R appuser:appuser credentials logs

# Switch to non-root user
USER appuser

# Set entrypoint
ENTRYPOINT ["python", "src/main.py"]
