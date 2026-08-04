FROM python:3.11-slim

WORKDIR /app

# Install dependencies first (cached layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app source
COPY . .

# HF Spaces Docker containers must expose port 7860
EXPOSE 7860

# Run Streamlit on port 7860 bound to all interfaces
# enableXsrfProtection / enableCORS are off because Spaces serves the app in an
# iframe behind a reverse proxy: the upload POST arrives cross-origin and the
# default XSRF check rejects it, so st.file_uploader silently fails.
CMD ["streamlit", "run", "app.py", \
     "--server.port=7860", \
     "--server.address=0.0.0.0", \
     "--server.headless=true", \
     "--server.fileWatcherType=none", \
     "--server.enableXsrfProtection=false", \
     "--server.enableCORS=false", \
     "--server.maxUploadSize=50", \
     "--browser.gatherUsageStats=false"]
