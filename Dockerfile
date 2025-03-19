# Use Python 3.9 slim image
FROM python:3.9-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first to leverage Docker cache
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Create necessary directories
RUN mkdir -p /app/uploads /app/pptx_folder /app/conversations

# Set environment variables
ENV PYTHONPATH=/app
ENV UPLOAD_FOLDER=/app/pptx_folder
ENV OUTPUT_FOLDER=/app/OUTPUT
ENV TEMPLATE_FILE=/app/templates/CRA_TEMPLATE_IA.pptx

# Expose port
EXPOSE 5050

# Run the application
CMD ["python", "src/api/api.py"] 