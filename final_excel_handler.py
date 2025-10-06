import asyncio
import logging
import time
from pathlib import Path
from typing import Dict, Any
import pythoncom
import win32com.client as win32

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class FinalExcelHandler:
    """Complete Excel handler - extracts ALL columns from Excel with CORRECT mappings"""
    
    def __init__(self, excel_path: str = "SmartOptionChainExcel_Zerodha.xlsm"):
        self.excel_path = Path(excel_path).resolve()
        self.SHEET_NAME = "Option_Chain"
        
    async def fetch_live_data(self, token_data: dict, symbol: str, 
                             option_expiry: str, future_expiry: str, 
                             chain_length: int) -> Dict[str, Any]:
        """Fetch complete live data with ALL columns"""
        
        try:
            pythoncom.CoInitialize()
            
            logger.info("Connecting to Excel...")
            excel = win32.GetActiveObject("Excel.Application")
            
            wb = None
            for workbook in excel.Workbooks:
                if self.excel_path.name in workbook.Name:
                    wb = workbook
                    break
            
            if not wb:
                raise Exception("Workbook not found")
            
            ws = wb.Sheets(self.SHEET_NAME)
            logger.info("Connected to sheet")
            
            # Write credentials
            logger.info("Writing credentials...")
            try:
                ws.Range("F587").Value = token_data['user_id']
                ws.Range("F615").Value = token_data.get('enc_token', token_data.get('enctoken', ''))
                logger.info("âœ“ Credentials written")
            except Exception as e:
                logger.error(f"Credential write failed: {e}")
            
            # Set inputs
            logger.info(f"Setting inputs: {symbol}, {option_expiry}")
            try:
                ws.Range("B2").Value = symbol
                ws.Range("B3").Value = option_expiry
                ws.Range("B4").Value = future_expiry
                ws.Range("B6").Value = chain_length
                logger.info("âœ“ Inputs set")
            except Exception as e:
                logger.error(f"Input setting failed: {e}")
            
            # Click button
            logger.info("Triggering data fetch...")
            button_clicked = False
            try:
                button = ws.Shapes("Button 2")
                if hasattr(button, 'OnAction') and button.OnAction:
                    excel.Run(button.OnAction)
                    button_clicked = True
                    logger.info("âœ“ Button clicked")
            except Exception as e:
                logger.warning(f"Button click failed: {e}")
            
            # Wait for refresh
            logger.info("Waiting 15 seconds for data refresh...")
            await asyncio.sleep(15)
            
            # Extract complete data
            logger.info("Extracting ALL data from Excel...")
            data = self._extract_complete_option_chain(ws, symbol, option_expiry, 
                                                      future_expiry, chain_length)
            
            logger.info(f"âœ“ Extracted {len(data.get('option_chain', []))} strikes")
            logger.info(f"âœ“ Spot LTP: {data.get('spot', {}).get('spot_ltp', 0)}")
            logger.info(f"âœ“ PCR: {data.get('pcr', 0)}")
            logger.info(f"âœ“ Max Pain: {data.get('max_pain', 0)}")
            
            return data
            
        except Exception as e:
            logger.error(f"FETCH FAILED: {e}")
            import traceback
            traceback.print_exc()
            
            return {
                'symbol': symbol,
                'option_expiry': option_expiry,
                'future_expiry': future_expiry,
                'error': str(e),
                'data_source': 'error'
            }
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _extract_complete_option_chain(self, ws, symbol, option_expiry, 
                                      future_expiry, chain_length) -> Dict[str, Any]:
        """Extract COMPLETE data - CORRECTED cell references based on your Excel"""
        
        try:
            # === TOP SECTION - Market Summary ===
            
            # OHLC Data (Column D, Rows 3-9)
            ohlc_data = {
                'open': self._safe_read_numeric(ws, "D3"),
                'high': self._safe_read_numeric(ws, "D4"),
                'low': self._safe_read_numeric(ws, "D5"),
                'close': self._safe_read_numeric(ws, "D6"),
                'ltp': self._safe_read_numeric(ws, "D7"),
                'ltp_change': self._safe_read_numeric(ws, "D8"),
                'ltp_change_pct': self._safe_read_numeric(ws, "D9")
            }
            
            # Spot Data (Column F, Rows 3-9)
            spot_data = {
                'spot': self._safe_read_numeric(ws, "F3"),
                'spot_ltp': self._safe_read_numeric(ws, "F7"),
                'spot_ltp_change': self._safe_read_numeric(ws, "F8"),
                'spot_ltp_change_pct': self._safe_read_numeric(ws, "F9")
            }
            
            # Future Data (Column G, Rows 3-9)
            future_data = {
                'future': self._safe_read_numeric(ws, "G3"),
                'future_price': self._safe_read_numeric(ws, "G7"),
                'future_change': self._safe_read_numeric(ws, "G8"),
                'future_change_pct': self._safe_read_numeric(ws, "G9")
            }
            
            # Open Interest Summary (Column I, Rows 4-9)
            open_interest_summary = {
                'open_interest': self._safe_read_numeric(ws, "I4"),
                'change_in_oi': self._safe_read_numeric(ws, "I5"),
                'max_oi': self._safe_read_numeric(ws, "I6"),
                'max_change_in_oi': self._safe_read_numeric(ws, "I7"),
                'max_oi_strike': self._safe_read_numeric(ws, "I8"),
                'max_change_in_oi_strike': self._safe_read_numeric(ws, "I9")
            }
            
            # Future OI (Column J)
            future_oi = {
                'future_oi': self._safe_read_numeric(ws, "J4"),
                'future_oi_change': self._safe_read_numeric(ws, "J5")
            }
            
            # Call Summary (Column K)
            call_summary = {
                'total_call_oi': self._safe_read_numeric(ws, "K4"),
                'total_call_volume': self._safe_read_numeric(ws, "K5"),
                'total_call_oi_change': self._safe_read_numeric(ws, "K6")
            }
            
            # Put Summary (Column L)
            put_summary = {
                'total_put_oi': self._safe_read_numeric(ws, "L4"),
                'total_put_volume': self._safe_read_numeric(ws, "L5"),
                'total_put_oi_change': self._safe_read_numeric(ws, "L6")
            }
            
            # Key Metrics - CORRECTED locations based on your specification
            pcr = self._safe_read_numeric(ws, "T2")
            max_pain = self._safe_read_numeric(ws, "T5")
            india_vix = self._safe_read_numeric(ws, "T8")
            
            # Trading Signals - CORRECTED locations
            intraday_signal = self._safe_read_text(ws, "P3")
            weekly_signal = self._safe_read_text(ws, "P7")
            
            # Build main data structure
            data = {
                'symbol': symbol,
                'option_expiry': option_expiry,
                'future_expiry': future_expiry,
                'chain_length': chain_length,
                'timestamp': time.strftime('%Y-%m-%d %H:%M:%S'),
                'data_source': 'excel_live',
                
                # Market Summary
                'ohlc': ohlc_data,
                'spot': spot_data,
                'future': future_data,
                'open_interest_summary': open_interest_summary,
                'future_open_interest': future_oi,
                'calls_summary': call_summary,
                'puts_summary': put_summary,
                
                # Key Metrics
                'pcr': pcr,
                'max_pain': max_pain,
                'india_vix': india_vix,
                
                # Signals
                'signals': {
                    'intraday': intraday_signal,
                    'weekly': weekly_signal
                },
                
                # Legacy compatibility
                'spot_ltp': spot_data['spot_ltp'],
                'market_data': {
                    'spot_ltp': spot_data['spot_ltp'],
                    'spot_ltp_change': spot_data['spot_ltp_change'],
                    'spot_ltp_change_pct': spot_data['spot_ltp_change_pct'],
                    'future_price': future_data['future_price'],
                    'pcr': pcr,
                    'max_pain': max_pain,
                    'india_vix': india_vix
                },
                
                'option_chain': []
            }
            
            # === OPTION CHAIN TABLE - CORRECTED COLUMN MAPPING ===
            # Data starts at row 13 (row 12 is header)
            start_row = 13
            
            logger.info(f"Extracting option chain with CORRECTED column mapping...")
            
            for i in range(chain_length):
                row = start_row + i
                
                # Strike (Column I) - CORRECTED: should be from I column, not row 11
                strike = self._safe_read_numeric(ws, f"I{row}")
                if not strike or strike < 1000:
                    continue
                
                # CORRECTED option data based on YOUR specification
                option_data = {
                    'strike': strike,
                    
                    # CALL data - CORRECTED columns (A-H)
                    'call': {
                        'interpretation': self._safe_read_text(ws, f"A{row}"),      # A12: INTERPRETATION
                        'avg_price': self._safe_read_numeric(ws, f"B{row}"),        # B12: Avg_Price
                        'iv': self._safe_read_numeric(ws, f"C{row}"),               # C12: IV
                        'oi_change': self._safe_read_numeric(ws, f"D{row}"),        # D12: OI_Change
                        'oi': self._safe_read_numeric(ws, f"E{row}"),               # E12: OI
                        'volume': self._safe_read_numeric(ws, f"F{row}"),           # F12: Volume
                        'ltp_change': self._safe_read_numeric(ws, f"G{row}"),       # G12: LTP_Change
                        'ltp': self._safe_read_numeric(ws, f"H{row}")               # H12: LTP
                    },
                    
                    # PUT data - CORRECTED columns (J-Q, in REVERSE order as per your spec)
                    'put': {
                        'ltp': self._safe_read_numeric(ws, f"J{row}"),              # J12: LTP
                        'ltp_change': self._safe_read_numeric(ws, f"K{row}"),       # K12: LTP_Change
                        'volume': self._safe_read_numeric(ws, f"L{row}"),           # L12: Volume
                        'oi': self._safe_read_numeric(ws, f"M{row}"),               # M12: OI
                        'oi_change': self._safe_read_numeric(ws, f"N{row}"),        # N12: OI_Change
                        'iv': self._safe_read_numeric(ws, f"O{row}"),               # O12: IV
                        'avg_price': self._safe_read_numeric(ws, f"P{row}"),        # P12: Avg_Price
                        'interpretation': self._safe_read_text(ws, f"Q{row}")       # Q12: INTERPRETATION
                    }
                }
                
                data['option_chain'].append(option_data)
            
            # Validation
            if spot_data['spot_ltp'] > 0 and len(data['option_chain']) > 0:
                logger.info("âœ“ Data validation: PASS")
            else:
                logger.warning("âš  Data validation: PARTIAL")
            
            logger.info(f"Summary:")
            logger.info(f"  Spot LTP: {spot_data['spot_ltp']}")
            logger.info(f"  PCR: {pcr}")
            logger.info(f"  Max Pain: {max_pain}")
            logger.info(f"  India VIX: {india_vix}")
            logger.info(f"  Intraday Signal: {intraday_signal}")
            logger.info(f"  Weekly Signal: {weekly_signal}")
            logger.info(f"  Strikes extracted: {len(data['option_chain'])}")
            
            # Show sample of first strike for verification
            if data['option_chain']:
                sample = data['option_chain'][0]
                logger.info(f"  Sample Strike {sample['strike']}:")
                logger.info(f"    Call LTP: {sample['call']['ltp']}, Put LTP: {sample['put']['ltp']}")
            
            return data
            
        except Exception as e:
            logger.error(f"Extraction error: {e}")
            import traceback
            traceback.print_exc()
            raise
    
    def _safe_read_numeric(self, ws, cell: str) -> float:
        """Safely read numeric value"""
        try:
            value = ws.Range(cell).Value
            
            if value is None or value == "":
                return 0.0
            
            if isinstance(value, (int, float)):
                return float(value)
            
            if isinstance(value, str):
                # Skip text headers
                skip_words = ['open', 'high', 'low', 'close', 'ltp', 'change', 
                             'interpretation', 'avg', 'price', 'volume', 'strike',
                             'calls', 'puts', 'oi', 'iv', 'pcr', 'pain', 'vix',
                             'bearish', 'bullish']
                
                if any(word in value.lower() for word in skip_words):
                    return 0.0
                
                try:
                    cleaned = value.replace(',', '').replace(' ', '').strip()
                    return float(cleaned)
                except:
                    return 0.0
            
            return 0.0
        except:
            return 0.0
    
    def _safe_read_text(self, ws, cell: str) -> str:
        """Safely read text value"""
        try:
            value = ws.Range(cell).Value
            if value is None or value == "":
                return ""
            return str(value).strip()
        except:
            return ""