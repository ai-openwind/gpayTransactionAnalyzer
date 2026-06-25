# GPay Transaction Analyzer

A small FastAPI application that extracts transaction data from Google Pay PDF statements and presents it through a browser interface or JSON/Excel APIs.

## Features

- Upload a Google Pay PDF statement via a web interface
- Parse transaction records and generate summaries
- Filter, search, and sort transactions in the browser
- Download parsed results as an Excel file
- JSON extraction endpoint for integration

## Project structure

- `main.py` - FastAPI application and PDF extraction API
- `extract_gpay.py` - standalone extraction script for local command-line use
- `static/index.html` - web UI for uploading and inspecting extracted transactions
- `requirements.txt` - Python dependencies
- `Dockerfile` - container image definition

## Requirements

- Python 3.11
- `pip`
- Google Pay PDF statement files

## Installation

1. Clone or open this repository.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Running locally

Start the server using Uvicorn:

```bash
uvicorn main:app --host 0.0.0.0 --port 8080
```

Then open a browser at `http://127.0.0.1:8080/`.

## Usage

### Web UI

- Drag and drop or select a Google Pay PDF file
- View extracted transactions and summary cards
- Filter by transaction type, date range, or search term
- Download parsed results as an Excel spreadsheet

### API endpoints

- `GET /` - serves the web UI
- `POST /extract` - upload a PDF and receive extracted data as JSON
- `POST /extract/excel` - upload a PDF and download an Excel workbook

Example JSON request with `curl`:

```bash
curl -F "file=@statement.pdf" http://127.0.0.1:8080/extract
```

Example Excel download request:

```bash
curl -o gpay_transactions.xlsx -F "file=@statement.pdf" http://127.0.0.1:8080/extract/excel
```

## Docker

Build and run using Docker:

```bash
docker build -t gpay-transaction-analyzer .
docker run -p 8080:8080 gpay-transaction-analyzer
```

The app listens on port `8080` by default.

## Notes

- The parser is designed for Google Pay PDF statements in the format used by the current template.
- If a PDF uses a different text layout, extraction may fail or miss some fields.
- The app supports `₹` currency formatting and basic UPI transaction detail extraction.

## License

This repository does not include a license file.

