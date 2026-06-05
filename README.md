# Elevator Packing List Translator

A small web app for uploading Chinese elevator packing list workbooks and generating:

- an English translated Excel workbook
- an A4 printable English PDF

## Local Run

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000`.

## Vercel Deploy

This repo is configured for Vercel Python serverless deployment.

```bash
vercel
```

For production:

```bash
vercel --prod
```

Uploaded files are processed in temporary serverless storage. The browser downloads a zip file containing the English PDF and Excel immediately after conversion.
