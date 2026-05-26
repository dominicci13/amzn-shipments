# amzn-shipments

Weekly ETL that pulls FBA shipment data from each Amazon Seller Central account, drops it as per-account CSVs that the connected workbook reads via Power Query, refreshes the workbook synchronously, and emails it. Runs **Tuesday 09:00 local** via APScheduler.

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
pip install git+https://github.com/dominicci13/shared-python-utils.git
```

### 2. Configure environment

```bash
cp .env.example .env
cp config/accounts.json.example config/accounts.json
cp config/paths.json.example config/paths.json
```

Fill in your credentials in `.env`, your Amazon account map in `accounts.json`, and your local OneDrive paths in `paths.json`.

## Run

```bash
python run_amzn_shipments.py
```

Prompts whether to run immediately, then schedules itself to run **Tue 09:00 local** via APScheduler.

## Project layout

```
amzn-shipments/
├── run_amzn_shipments.py       # Single-file entry — scrape + refresh + email
├── config/
│   ├── accounts.json           # Amazon account names + Seller Central URLs (gitignored)
│   ├── accounts.json.example   # Template
│   ├── paths.json              # Workbook + downloads paths (gitignored)
│   └── paths.json.example      # Template
├── vba/
│   └── modUtilities.bas        # Version-controlled VBA — synchronous `refresh` sub
├── logs/                       # Rotating log files (gitignored)
├── screenshots/                # Debug screenshots written on browser errors (gitignored)
├── output/                     # Future use; currently empty (gitignored)
├── requirements.txt
├── LICENSE
└── README.md
```

## Environment Variables

| Variable | Description |
|---|---|
| `AMZN_email` | Amazon Seller Central login email |
| `AMZN_pass` | Amazon Seller Central password |
| `CHROME_USER_DATA_DIR` | Path to Chrome automation profile directory |
| `ALERT_EMAIL` | Outlook recipient for unhandled-exception crash reports |
| `SENDER_EMAIL` | Outlook account used to send the report email |
| `TO_EMAIL` | Comma-separated list of recipient email addresses |
| `CC_EMAIL` | Optional comma-separated list of CC email addresses |

## Author

Built by **Brian Ramirez** ([@dominicci13](https://github.com/dominicci13)) — automation & AI workflow specialist. More on my [GitHub profile](https://github.com/dominicci13) and [LinkedIn](https://linkedin.com/in/bdramirez).

## License

[MIT](LICENSE)
