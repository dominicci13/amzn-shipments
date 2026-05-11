# amzn-shipments

Downloads and processes Amazon FBA shipment data from Seller Central, updates a SQL Server database, and emails a weekly summary. Runs on a Tuesday schedule via APScheduler.

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
pip install git+https://github.com/dominicci13/shared-python-utils.git
```

### 2. Configure environment

```bash
cp .env.example .env
```

Edit `.env` with your credentials.

## Run

```bash
python run_amzn_shipments.py
```

The script runs automatically at 09:00 on Tuesdays via APScheduler.

## Environment Variables

| Variable | Description |
|---|---|
| `AMZN_email` | Amazon Seller Central login email |
| `AMZN_pass` | Amazon Seller Central password |
| `CHROME_USER_DATA_DIR` | Path to Chrome automation profile directory |
| `SENDER_EMAIL` | Outlook account used to send the report email |
| `TO_EMAIL` | Comma-separated list of recipient email addresses |
| `CC_EMAIL` | Comma-separated list of CC email addresses |
