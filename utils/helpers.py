import logging
from datetime import datetime, time
from pathlib import Path

def setup_logging(log_file: Path, level: str = "INFO"):
    """Setup logging configuration"""
    log_file.parent.mkdir(exist_ok=True)
    
    logging.basicConfig(
        level=getattr(logging, level),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )

def is_market_hours() -> bool:
    """Check if current time is within market hours"""
    now = datetime.now().time()
    start_time = time(9, 15)  # 9:15 AM
    end_time = time(15, 30)   # 3:30 PM
    return start_time <= now <= end_time

def format_currency(value: float) -> str:
    """Format currency values"""
    if value is None:
        return "N/A"
    return f"₹{value:,.2f}"

def format_number(value: int) -> str:
    """Format number with commas"""
    if value is None:
        return "0"
    return f"{value:,}"