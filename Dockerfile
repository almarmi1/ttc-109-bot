FROM python:3.11-slim

# Prevent Python from buffering stdout/stderr so logs show immediately in Coolify
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Run your Python bot script (replace main.py with your entry file)
CMD ["python", "main.py"]
