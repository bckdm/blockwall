FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn
COPY app.py ./
COPY app ./app
COPY data /app/seed
COPY entrypoint.sh /app/entrypoint.sh
RUN chmod +x /app/entrypoint.sh
RUN mkdir -p /app/data
EXPOSE 5001
ENTRYPOINT ["/app/entrypoint.sh"]
CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:5001", "app:app"]