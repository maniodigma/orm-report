# Drop & Go Web App (Streamlit)

**For ORM team (non-technical):** Upload the monthly XLSX and (optionally) the logo → click download. You’ll get both a **CEO-ready PPTX** and an **interactive HTML deck**.

## Run locally (easy)
```bash
pip install -r requirements.txt
streamlit run app.py
```
Open the local URL shown in the terminal (typically http://localhost:8501).

## Deploy options
### A) Streamlit Community Cloud (free)
1. Push this folder to a GitHub repo.
2. Go to https://share.streamlit.io → “New app” → select your repo → main branch → app file: `app.py`.
3. App launches with a public link you can share.

### B) Internal server
```bash
pip install -r requirements.txt
streamlit run app.py --server.port 8080 --server.address 0.0.0.0
```
Reverse-proxy behind Nginx if needed.

### C) Docker (optional)
```bash
# From this folder
cat > Dockerfile <<'EOF'
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
EOF

docker build -t sumadhura-report:latest .
docker run -p 8501:8501 sumadhura-report:latest
```

## Notes
- **Header row:** defaults to `14` (A14). Change in the sidebar.
- **Date column:** defaults to `Date Reported`. Adjust in the sidebar if needed.
- Pie slices under the threshold are grouped into **“Other”** automatically.
- All branding (colors and logo) is in the left sidebar; set it once and reuse.
