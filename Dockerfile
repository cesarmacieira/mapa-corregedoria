# syntax=docker/dockerfile:1
FROM python:3.11-slim
 
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    STREAMLIT_SERVER_PORT=8501 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
    MPLBACKEND=Agg
 
WORKDIR /app
 
COPY requirements.txt .
RUN pip install --upgrade pip && pip install --no-cache-dir -r requirements.txt
 
COPY app.py .
COPY SERVIDORES_ADM_15_07_2025.csv .
COPY SERVIDORES_JUD_15_07_2025.csv .
 
RUN useradd -m -u 1000 -s /bin/bash appuser && chown -R appuser:appuser /app
USER appuser
ENV HOME=/home/appuser
 
EXPOSE 8501
 
CMD ["streamlit", "run", "app.py"]