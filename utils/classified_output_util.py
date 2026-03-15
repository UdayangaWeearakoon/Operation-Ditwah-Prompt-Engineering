from pathlib import Path
from datetime import datetime
import pandas as pd

OUTPUT_PATH = Path("..\\output\\classified_messages.xlsx")

HEADERS = [
    "timestamp",
    "district",
    "intent",
    "priority",
]


def _ensure_output_file() -> None:
    """
    Create output directory and Excel file with headers if not exists.
    """
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not OUTPUT_PATH.exists():
        df = pd.DataFrame(columns=HEADERS)
        df.to_excel(OUTPUT_PATH, index=False, engine="openpyxl")


def save_classified_message(
    district: str,
    intent: str,
    priority: str
) -> None:
    """
    Save classified message output to Excel.

    Format:
    District: [Name] | Intent: [Category] | Priority: [High/Low]
    """

    _ensure_output_file()

    # Normalize values
    priority = priority.capitalize()
    intent = intent.strip()
    district = district.strip()

    if priority not in {"High", "Low"}:
        raise ValueError("Priority must be 'High' or 'Low'")

    new_row = {
        "timestamp": datetime.now().isoformat(),
        "district": district,
        "intent": intent,
        "priority": priority,
    }

    df = pd.read_excel(OUTPUT_PATH, engine="openpyxl")
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    df.to_excel(OUTPUT_PATH, index=False, engine="openpyxl")