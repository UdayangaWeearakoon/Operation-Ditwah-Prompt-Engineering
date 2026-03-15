from pathlib import Path
from datetime import datetime
import pandas as pd
import os

OUTPUT_PATH = Path("..") / "output" / "flood_report.xlsx"

HEADERS = [
    "timestamp",
    "district",
    "flood_level_meters",
    "victim_count",
    "main_need",
    "status",
]

def _ensure_output_file() -> None:
    """Creates the directory and the Excel file with headers if it doesn't exist."""
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    
    if not OUTPUT_PATH.exists():
        df = pd.DataFrame(columns=HEADERS)
        df.to_excel(OUTPUT_PATH, index=False, engine="openpyxl")

def save_events_to_excel(events: list) -> None:
    """
    Appends flood events to the Excel report.
    Accepts a list of dictionaries or Pydantic models.
    """
    if not events:
        return

    _ensure_output_file()

    processed_data = []
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for event in events:
        # If the input is a Pydantic model, convert to dict
        # Otherwise, assume it's already a dictionary
        data = event if isinstance(event, dict) else event.dict()

        # Clean and normalize data
        normalized_event = {
            "timestamp": current_time,
            "district": str(data.get("district", "Unknown")).strip(),
            "flood_level_meters": float(data.get("flood_level_meters") or 0.0),
            "victim_count": int(data.get("victim_count") or 0),
            "main_need": str(data.get("main_need", "N/A")).strip(),
            "status": str(data.get("status", "Stable")).strip(),
        }
        processed_data.append(normalized_event)

    # Efficiently append to existing data
    try:
        existing_df = pd.read_excel(OUTPUT_PATH, engine="openpyxl")
        new_df = pd.DataFrame(processed_data)
        
        # Combine and save
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
        updated_df.to_excel(OUTPUT_PATH, index=False, engine="openpyxl")
        print(f"Successfully saved {len(processed_data)} events to {OUTPUT_PATH}")
        
    except Exception as e:
        print(f"Error writing to Excel: {e}")