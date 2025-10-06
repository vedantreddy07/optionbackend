import os
from pathlib import Path

# Project paths
PROJECT_ROOT = Path(__file__).parent.parent
EXE_PATH = PROJECT_ROOT / "SmartOptionChainExcel.exe"
EXCEL_PATH = PROJECT_ROOT / "SmartOptionChainExcel_Zerodha.xlsm"

# Market settings
MARKET_START_TIME = "09:15"
MARKET_END_TIME = "15:30"

# Excel cell mappings (adjust these based on your actual Excel file)
EXCEL_CELLS = {
    "USER_ID": "R27",
    "ENC_TOKEN": "R29",
    "SYMBOL": "B2",
    "OPTION_EXPIRY": "B3",
    "FUTURE_EXPIRY": "B4",
    "CHAIN_LENGTH": "B6",
    "UNDERLYING_PRICE": "F2",
    "DATA_START_ROW": 10
}

# Logging configuration
LOG_LEVEL = "INFO"
LOG_FILE = PROJECT_ROOT / "logs" / "system.log"