import json
import os
import uuid
import pandas as pd
from datetime import date, timedelta

DATA_FILE = os.path.join(os.path.dirname(__file__), '..', 'data', 'spend_log.json')

def _ensure_data_file():
    """Create data file if it doesn't exist."""
    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    if not os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'w') as f:
            json.dump([], f)

def load_data() -> pd.DataFrame:
    """Load all entries as a DataFrame."""
    _ensure_data_file()
    with open(DATA_FILE, 'r') as f:
        data = json.load(f)
    if not data:
        return pd.DataFrame(columns=[
            'id', 'week_ending', 'provider_name', 'provider_type',
            'specialty', 'service_line', 'department',
            'days_worked', 'daily_rate', 'total_spend', 'notes', 'logged_at'
        ])
    df = pd.DataFrame(data)
    df['days_worked'] = pd.to_numeric(df['days_worked'], errors='coerce')
    df['daily_rate'] = pd.to_numeric(df['daily_rate'], errors='coerce')
    df['total_spend'] = pd.to_numeric(df['total_spend'], errors='coerce')
    return df

def save_entry(entry: dict):
    """Append a new entry with a unique ID."""
    _ensure_data_file()
    with open(DATA_FILE, 'r') as f:
        data = json.load(f)
    entry['id'] = str(uuid.uuid4())
    data.append(entry)
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f, indent=2)

def delete_entry(entry_id: str):
    """Delete an entry by ID."""
    _ensure_data_file()
    with open(DATA_FILE, 'r') as f:
        data = json.load(f)
    data = [e for e in data if e.get('id') != entry_id]
    with open(DATA_FILE, 'w') as f:
        json.dump(data, f, indent=2)

def get_week_ending() -> date:
    """Return the upcoming or current Saturday as the default week ending date."""
    today = date.today()
    days_until_saturday = (5 - today.weekday()) % 7
    if days_until_saturday == 0:
        return today
    return today + timedelta(days=days_until_saturday)
