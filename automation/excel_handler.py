import asyncio
import logging
import time
from pathlib import Path
from typing import Dict, Any, List
import pythoncom
import win32com.client as win32

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class FinalExcelHandler:
    """Excel handler that clicks OPTION CHAIN button to trigger refresh"""
    
    def __init__(self, excel_path: str = "SmartOptionChainExcel_Zerodha.xlsm"):
        self.excel_path = Path(excel_path).resolve()
        self.SHEET_NAME = "Option_Chain"
        
    async def fetch_live_data(self, token_data: dict, symbol: str, 
                             option_expiry: str, future_expiry: str, 
                             chain_length: int) -> Dict[str, Any]:
        """Complete workflow: credentials → inputs → click button → extract data"""
        
        try:
            pythoncom.CoInitialize()
            
            # Connect to Excel
            logger.info("Connecting to Excel...")
            try:
                excel = win32.GetActiveObject("Excel.Application")
            except:
                excel = win32.Dispatch("Excel.Application")
                excel.Visible = True
            
            # Find workbook
            wb = None
            for workbook in excel.Workbooks:
                if self.excel_path.name in workbook.Name:
                    wb = workbook
                    break
            
            if not wb:
                wb = excel.Workbooks.Open(str(self.excel_path))
            
            ws = wb.Sheets(self.SHEET_NAME)
            logger.info(f"Connected to sheet: {self.SHEET_NAME}")
            
            # Step 1: Enter credentials
            logger.info("Entering credentials...")
            ws.Range("F587").Value = token_data['user_id']
            ws.Range("F615").Value = token_data['enc_token']
            
            # Step 2: Set inputs
            logger.info(f"Setting inputs: {symbol}, {option_expiry}, length={chain_length}")
            ws.Range("B2").Value = symbol
            ws.Range("B3").Value = option_expiry
            ws.Range("B4").Value = future_expiry
            ws.Range("B6").Value = chain_length
            
            # Step 3: Click "OPTION CHAIN" button (THIS IS THE KEY!)
            logger.info("Clicking OPTION CHAIN button to trigger data fetch...")
            try:
                # Method 1: Try direct OnAction
                button = ws.Shapes("Button 2")
                if hasattr(button, 'OnAction') and button.OnAction:
                    logger.info(f"Executing OnAction: {button.OnAction}")
                    excel.Run(button.OnAction)
                else:
                    # Method 2: Select the button (simulates click)
                    logger.info("Selecting button to simulate click")
                    button.Select()
                    
            except Exception as e:
                logger.error(f"Button click failed: {e}")
                # Try alternative: use SendKeys to press button
                try:
                    excel.Activate()
                    wb.Activate()
                    ws.Activate()
                    import win32api, win32con
                    # Tab to button and press Space/Enter
                    for _ in range(3):
                        win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
                        win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
                        time.sleep(0.2)
                    win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                    win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                except Exception as e2:
                    logger.error(f"Alternative click method failed: {e2}")
            
            # Step 4: Wait for data to populate
            logger.info("Waiting for Zerodha API to fetch data (20 seconds)...")
            await asyncio.sleep(20)
            
            # Step 5: Extract complete data
            logger.info("Extracting data from Excel...")
            data = self._extract_complete_option_chain(ws, symbol, option_expiry, 
                                                      future_expiry, chain_length)
            
            logger.info(f"Extracted {len(data.get('option_chain', []))} strikes")
            logger.info(f"Spot LTP: {data.get('spot_ltp')}")
            
            return data
            
        except Exception as e:
            logger.error(f"Data fetch failed: {e}")
            raise
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _extract_complete_option_chain(self, ws, symbol, option_expiry, 
                                      future_expiry, chain_length) -> Dict[str, Any]:
        """Extract all data matching Excel structure"""
        
        try:
            # Basic market data (corrected cell references)
            data = {
                'symbol': symbol,
                'option_expiry': option_expiry,
                'future_expiry': future_expiry,
                'chain_length': chain_length,
                'timestamp': time.strftime('%Y-%m-%d %H:%M:%S'),
                
                # Spot data - CORRECTED: F6 not D6
                'spot_ltp': self._safe_read(ws, "F6"),
                'spot_ltp_change': self._safe_read(ws, "F7"),
                'spot_ltp_change_pct': self._safe_read(ws, "F8"),
                
                # Future price
                'future_price': self._safe_read(ws, "G2"),
                
                # Market indicators
                'pcr': self._safe_read(ws, "Q2"),
                'max_pain': self._safe_read(ws, "Q4"),
                'india_vix': self._safe_read(ws, "Q6"),
                
                'data_source': 'excel_live'
            }
            
            # Extract option chain (rows 13 onwards)
            option_chain = []
            start_row = 13
            
            for i in range(chain_length):
                row = start_row + i
                
                strike = self._safe_read(ws, f"J{row}")
                if not strike or strike == 0:
                    continue
                
                option_data = {
                    'strike': strike,
                    'call': {
                        'ltp': self._safe_read(ws, f"I{row}"),
                        'ltp_change': self._safe_read(ws, f"H{row}"),
                        'volume': self._safe_read(ws, f"G{row}"),
                        'oi': self._safe_read(ws, f"F{row}"),
                        'oi_change': self._safe_read(ws, f"E{row}"),
                        'iv': self._safe_read(ws, f"D{row}"),
                        'avg_price': self._safe_read(ws, f"C{row}"),
                        'interpretation': self._safe_read(ws, f"B{row}") or ""
                    },
                    'put': {
                        'ltp': self._safe_read(ws, f"K{row}"),
                        'ltp_change': self._safe_read(ws, f"L{row}"),
                        'volume': self._safe_read(ws, f"M{row}"),
                        'oi': self._safe_read(ws, f"N{row}"),
                        'oi_change': self._safe_read(ws, f"O{row}"),
                        'iv': self._safe_read(ws, f"P{row}"),
                        'avg_price': self._safe_read(ws, f"Q{row}"),
                        'interpretation': self._safe_read(ws, f"R{row}") or ""
                    }
                }
                
                option_chain.append(option_data)
            
            data['option_chain'] = option_chain
            
            # Validate data
            if isinstance(data['spot_ltp'], (int, float)) and data['spot_ltp'] > 0:
                logger.info("Data validation: PASS")
            else:
                logger.warning(f"Data validation: FAIL - spot_ltp={data['spot_ltp']}")
                data['data_source'] = 'excel_stale'
            
            return data
            
        except Exception as e:
            logger.error(f"Data extraction error: {e}")
            raise
    
    def _safe_read(self, ws, cell: str):
        """Safely read cell value with type conversion"""
        try:
            value = ws.Range(cell).Value
            
            # Handle None/Empty
            if value is None or value == "":
                return 0
            
            # Handle text labels
            if isinstance(value, str):
                if value.lower() in ['close', 'open', 'high', 'low', 'ltp']:
                    return 0
                # Try to parse as number
                try:
                    return float(value.replace(',', ''))
                except:
                    return value
            
            return value
            
        except Exception as e:
            logger.debug(f"Error reading {cell}: {e}")
            return 0