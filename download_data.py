import os
import requests
import zipfile
import io

os.makedirs("data", exist_ok=True)
os.makedirs("charts", exist_ok=True)

excel_path = os.path.join("data", "Online Retail.xlsx")

if os.path.exists(excel_path):
    print(f"Dataset already exists at {excel_path} ({os.path.getsize(excel_path)} bytes).")
else:
    url = "https://archive.ics.uci.edu/static/public/352/online+retail.zip"
    print(f"Downloading UCI dataset from {url}...")
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    resp = requests.get(url, headers=headers, stream=True, timeout=120)
    resp.raise_for_status()
    
    zip_bytes = io.BytesIO()
    downloaded = 0
    for chunk in resp.iter_content(chunk_size=1024 * 1024):
        if chunk:
            zip_bytes.write(chunk)
            downloaded += len(chunk)
            print(f"Downloaded {downloaded / (1024 * 1024):.2f} MB...")
            
    print("Download complete. Extracting 'Online Retail.xlsx'...")
    zip_bytes.seek(0)
    with zipfile.ZipFile(zip_bytes) as z:
        z.extract("Online Retail.xlsx", path="data")
        
    print(f"Extracted successfully: {excel_path} ({os.path.getsize(excel_path):,} bytes)")
