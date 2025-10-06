import logging
import sys
from typing import Optional
from datetime import datetime

# Optional Windows COM imports for cross-platform compatibility
try:
    import pythoncom  # type: ignore
    import win32com.client as win32  # type: ignore
    WIN32_AVAILABLE = (sys.platform == "win32")
except Exception:
    pythoncom = None  # type: ignore
    win32 = None  # type: ignore
    WIN32_AVAILABLE = False
from pathlib import Path
from typing import Dict, List
import json
from datetime import datetime, timedelta

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class EnhancedDateExtractor:
    """Enhanced Excel date extractor with multiple fallback methods"""
    
    def __init__(self, excel_path: str = "SmartOptionChainExcel_Zerodha.xlsm"):
        self.excel_path = Path(excel_path).resolve()
        self.SHEET_NAME = "Option_Chain"
    
    def _format_date_string(self, date_value) -> str:
        """Convert any date format to DD-MM-YYYY string"""
        try:
            if date_value is None:
                return None
            
            # If it's already a string
            if isinstance(date_value, str):
                date_str = date_value.strip()
                
                # Remove timezone info
                date_str = date_str.split('+')[0].split('.')[0].strip()
                
                # Try to parse and reformat
                for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y']:
                    try:
                        dt = datetime.strptime(date_str, fmt)
                        return dt.strftime('%d-%m-%Y')
                    except ValueError:
                        continue
                
                # If it's already in DD-MM-YYYY or DD/MM/YYYY format
                if self._looks_like_date(date_str):
                    return date_str.replace('/', '-')
                
                return date_str
            
            # If it's a datetime object
            if hasattr(date_value, 'strftime'):
                return date_value.strftime('%d-%m-%Y')
            
            # If it's a COM date/time
            if isinstance(date_value, (int, float)):
                try:
                    # Excel date serial number
                    dt = datetime(1899, 12, 30) + timedelta(days=date_value)
                    return dt.strftime('%d-%m-%Y')
                except:
                    pass
            
            return str(date_value)
            
        except Exception as e:
            logger.debug(f"Date formatting error: {e}")
            return str(date_value) if date_value else None
    
    def extract_all_dates(self) -> Dict[str, List[str]]:
        """Extract dates using multiple methods with fallbacks"""
        try:
            if not WIN32_AVAILABLE:
                logger.warning("pywin32/pythoncom not available or non-Windows OS. Using fallback dates.")
                return self._get_fallback_dates()

            pythoncom.CoInitialize()
            
            logger.info("="*70)
            logger.info("EXTRACTING DATES FROM EXCEL")
            logger.info("="*70)
            
            # Connect to Excel
            excel = self._connect_to_excel()
            if not excel:
                return self._get_fallback_dates()
            
            wb = self._find_workbook(excel)
            if not wb:
                return self._get_fallback_dates()
            
            ws = wb.Sheets(self.SHEET_NAME)
            logger.info(f"Connected to sheet: {self.SHEET_NAME}")
            
            # Try multiple extraction methods
            results = {
                'symbols': [],
                'option_expiry': [],
                'future_expiry': []
            }
            
            # Method 1: Data Validation
            logger.info("\n[Method 1] Trying data validation extraction...")
            symbols = self._extract_from_validation(ws, "B2")
            option_dates = self._extract_from_validation(ws, "B3")
            future_dates = self._extract_from_validation(ws, "B4")
            
            if symbols:
                logger.info(f"  ✓ Symbols from validation: {len(symbols)} items")
                results['symbols'] = symbols
            if option_dates:
                # Format all dates
                option_dates = [self._format_date_string(d) for d in option_dates if d]
                option_dates = [d for d in option_dates if d]  # Remove None values
                logger.info(f"  ✓ Option dates from validation: {len(option_dates)} items")
                results['option_expiry'] = option_dates
            if future_dates:
                # Format all dates
                future_dates = [self._format_date_string(d) for d in future_dates if d]
                future_dates = [d for d in future_dates if d]  # Remove None values
                logger.info(f"  ✓ Future dates from validation: {len(future_dates)} items")
                results['future_expiry'] = future_dates
            
            # Method 2: Named Ranges
            if not results['symbols']:
                logger.info("\n[Method 2] Trying named ranges...")
                results['symbols'] = self._extract_from_named_range(wb, ws, "SymbolList")
            
            if not results['option_expiry']:
                option_dates = self._extract_from_named_range(wb, ws, "OptionExpiryList")
                results['option_expiry'] = [self._format_date_string(d) for d in option_dates if d]
            
            if not results['future_expiry']:
                future_dates = self._extract_from_named_range(wb, ws, "FutureExpiryList")
                results['future_expiry'] = [self._format_date_string(d) for d in future_dates if d]
            
            # Method 3: Scan for date-like values in columns
            if not results['option_expiry']:
                logger.info("\n[Method 3] Scanning for date patterns in sheet...")
                scanned_dates = self._scan_for_dates(ws, max_rows=200)
                results['option_expiry'] = [self._format_date_string(d) for d in scanned_dates if d]
            
            if not results['future_expiry'] and results['option_expiry']:
                # Use option expiry dates as future expiry if not found
                results['future_expiry'] = results['option_expiry'].copy()
            
            # Method 4: Check dropdown cells' current values and nearby cells
            if not results['option_expiry']:
                logger.info("\n[Method 4] Checking dropdown cell values...")
                current_option = ws.Range("B3").Value
                current_future = ws.Range("B4").Value
                
                if current_option:
                    formatted = self._format_date_string(current_option)
                    if formatted and self._looks_like_date(formatted):
                        results['option_expiry'] = [formatted]
                        logger.info(f"  ✓ Found current option expiry: {formatted}")
                
                if current_future:
                    formatted = self._format_date_string(current_future)
                    if formatted and self._looks_like_date(formatted):
                        results['future_expiry'] = [formatted]
                        logger.info(f"  ✓ Found current future expiry: {formatted}")
            
            # Validate and clean results
            results = self._validate_and_clean(results)
            
            # Save results
            self._save_to_cache(results)
            
            # Print summary
            self._print_summary(results)
            
            return results
            
        except Exception as e:
            logger.error(f"Date extraction failed: {e}")
            import traceback
            traceback.print_exc()
            return self._get_fallback_dates()
        
        finally:
            try:
                if WIN32_AVAILABLE and pythoncom is not None:
                    pythoncom.CoUninitialize()
            except:
                pass
    
    def _connect_to_excel(self):
        """Connect to running Excel instance"""
        try:
            excel = win32.GetActiveObject("Excel.Application")
            logger.info("✓ Connected to running Excel instance")
            return excel
        except:
            try:
                excel = win32.Dispatch("Excel.Application")
                excel.Visible = False
                logger.info("✓ Created new Excel instance")
                return excel
            except Exception as e:
                logger.error(f"✗ Cannot connect to Excel: {e}")
                return None
    
    def _find_workbook(self, excel):
        """Find the target workbook"""
        try:
            for wb in excel.Workbooks:
                if self.excel_path.name in wb.Name or "SmartOptionChain" in wb.Name:
                    logger.info(f"✓ Found workbook: {wb.Name}")
                    return wb
            
            # Try to open
            logger.info(f"Opening workbook: {self.excel_path}")
            return excel.Workbooks.Open(str(self.excel_path))
        except Exception as e:
            logger.error(f"✗ Cannot find/open workbook: {e}")
            return None
    
    def _extract_from_validation(self, ws, cell_ref: str) -> List[str]:
        """Extract from data validation"""
        try:
            cell = ws.Range(cell_ref)
            
            if not hasattr(cell, 'Validation'):
                return []
            
            validation = cell.Validation
            
            if validation.Type != 3:  # Not a list validation
                return []
            
            formula = validation.Formula1
            logger.info(f"  Validation formula for {cell_ref}: {formula}")
            
            # Parse formula
            if formula.startswith("="):
                # Range reference
                range_ref = formula[1:]
                
                try:
                    # Try to parse if it's a different sheet reference
                    if "!" in range_ref:
                        # Format: SheetName!$A$1:$A$10
                        sheet_name, cell_range = range_ref.split("!")
                        logger.info(f"    Reading from sheet: {sheet_name}, range: {cell_range}")
                        
                        # Get the workbook
                        wb = ws.Parent
                        
                        # Get the referenced sheet
                        try:
                            target_sheet = wb.Sheets(sheet_name)
                        except:
                            logger.warning(f"    Sheet '{sheet_name}' not found")
                            return []
                        
                        # Read the range from the target sheet
                        values = []
                        range_data = target_sheet.Range(cell_range).Value
                        
                        if isinstance(range_data, tuple):
                            # Multiple cells
                            for item in range_data:
                                if isinstance(item, tuple):
                                    val = item[0]
                                else:
                                    val = item
                                
                                if val and str(val).strip() and str(val) != "None":
                                    values.append(val)  # Keep original format, will format later
                        elif range_data:
                            # Single cell
                            values = [range_data]
                        
                        logger.info(f"    Extracted {len(values)} items from {sheet_name}")
                        return values
                    
                    else:
                        # Same sheet reference
                        range_obj = ws.Parent.Evaluate(range_ref)
                        
                        if isinstance(range_obj, tuple):
                            values = [item[0] for item in range_obj if item and item[0]]
                        else:
                            values = [range_obj] if range_obj else []
                        
                        return [v for v in values if v and str(v) != "None"]
                
                except Exception as e:
                    logger.warning(f"  Range extraction failed: {e}")
                    import traceback
                    traceback.print_exc()
            else:
                # Direct list
                return [v.strip() for v in formula.split(",") if v.strip()]
        
        except Exception as e:
            logger.debug(f"Validation extraction failed for {cell_ref}: {e}")
        
        return []
    
    def _extract_from_named_range(self, wb, ws, range_name: str) -> List[str]:
        """Extract from named range"""
        try:
            named_range = wb.Names(range_name)
            range_ref = named_range.RefersToRange
            
            values = []
            for cell in range_ref:
                if cell.Value:
                    values.append(cell.Value)
            
            if values:
                logger.info(f"  ✓ Found {len(values)} items in named range '{range_name}'")
            return values
        
        except Exception as e:
            logger.debug(f"Named range '{range_name}' not found: {e}")
            return []
    
    def _scan_for_dates(self, ws, max_rows: int = 200) -> List[str]:
        """Scan sheet for date patterns"""
        dates_found = set()
        
        for col in range(1, 31):  # Scan first 30 columns
            for row in range(1, max_rows):
                try:
                    value = ws.Cells(row, col).Value
                    if value:
                        str_val = str(value)
                        if self._looks_like_date(str_val):
                            dates_found.add(value)
                            if len(dates_found) >= 20:
                                break
                except:
                    continue
            if len(dates_found) >= 20:
                break
        
        dates_list = sorted(list(dates_found), key=lambda x: self._parse_date_for_sort(x))
        if dates_list:
            logger.info(f"  ✓ Found {len(dates_list)} dates by scanning")
        return dates_list
    
    def _parse_date_for_sort(self, date_val):
        """Parse date for sorting"""
        try:
            formatted = self._format_date_string(date_val)
            if formatted:
                return datetime.strptime(formatted, '%d-%m-%Y')
        except:
            pass
        return datetime(1900, 1, 1)
    
    def _looks_like_date(self, value: str) -> bool:
        """Check if string looks like a date"""
        if not value or len(str(value)) < 5:
            return False
        
        value = str(value)
        
        # DD-MM-YYYY or DD/MM/YYYY
        if ('-' in value or '/' in value) and any(c.isdigit() for c in value):
            parts = value.replace('/', '-').split('-')
            if len(parts) == 3:
                try:
                    day, month, year = parts
                    return (day.isdigit() and month.isdigit() and year.isdigit() and
                            1 <= int(day) <= 31 and 1 <= int(month) <= 12)
                except:
                    pass
        
        return False
    
    def _validate_and_clean(self, results: Dict[str, List[str]]) -> Dict[str, List[str]]:
        """Validate and clean extracted data"""
        
        # Remove duplicates while preserving order
        for key in results:
            results[key] = list(dict.fromkeys(results[key]))
        
        # Remove None and empty strings
        for key in results:
            results[key] = [v for v in results[key] if v and str(v).strip() and str(v) != "None"]
        
        # Use fallback if still empty
        if not results['symbols']:
            results['symbols'] = ["NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY", "SENSEX"]
            logger.warning("  ⚠ Using fallback symbols")
        
        if not results['option_expiry']:
            logger.warning("  ⚠ No option expiry dates found, using fallback")
            results['option_expiry'] = self._generate_fallback_dates()
        
        if not results['future_expiry']:
            logger.warning("  ⚠ Using option dates as future dates")
            results['future_expiry'] = results['option_expiry'].copy()
        
        return results
    
    def _generate_fallback_dates(self) -> List[str]:
        """Generate fallback dates (next 12 weeks)"""
        from datetime import timedelta
        
        dates = []
        current = datetime.now()
        
        for week in range(12):
            # Find next Thursday
            days_ahead = (3 - current.weekday()) % 7
            if days_ahead == 0:
                days_ahead = 7
            
            thursday = current + timedelta(days=days_ahead)
            dates.append(thursday.strftime("%d-%m-%Y"))
            current = thursday + timedelta(days=7)
        
        return dates
    
    def _save_to_cache(self, results: Dict):
        """Save to cache file"""
        try:
            cache_file = Path("excel_dates_cache.json")
            with open(cache_file, 'w') as f:
                json.dump({
                    **results,
                    'extracted_at': datetime.now().isoformat(),
                    'source': 'excel_extraction'
                }, f, indent=2)
            logger.info(f"\n✓ Dates cached to {cache_file}")
        except Exception as e:
            logger.warning(f"Failed to cache dates: {e}")
    
    def _print_summary(self, results: Dict):
        """Print extraction summary"""
        logger.info("\n" + "="*70)
        logger.info("EXTRACTION SUMMARY")
        logger.info("="*70)
        
        logger.info(f"\n✓ SYMBOLS ({len(results['symbols'])}):")
        for sym in results['symbols'][:10]:
            logger.info(f"    {sym}")
        if len(results['symbols']) > 10:
            logger.info(f"    ... and {len(results['symbols']) - 10} more")
        
        logger.info(f"\n✓ OPTION EXPIRY DATES ({len(results['option_expiry'])}):")
        for i, date in enumerate(results['option_expiry'][:10], 1):
            logger.info(f"    {i:2d}. {date}")
        if len(results['option_expiry']) > 10:
            logger.info(f"    ... and {len(results['option_expiry']) - 10} more")
        
        logger.info(f"\n✓ FUTURE EXPIRY DATES ({len(results['future_expiry'])}):")
        for i, date in enumerate(results['future_expiry'], 1):
            logger.info(f"    {i:2d}. {date}")
        
        logger.info("="*70)
    
    def _get_fallback_dates(self) -> Dict[str, List[str]]:
        """Fallback dates if extraction completely fails"""
        logger.warning("\n⚠ Using complete fallback dates")
        return {
            'symbols': ["NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY", "SENSEX"],
            'option_expiry': self._generate_fallback_dates(),
            'future_expiry': self._generate_fallback_dates()
        }


if __name__ == "__main__":
    print("\n" + "="*70)
    print("EXCEL DATE EXTRACTOR - ENHANCED VERSION")
    print("="*70)
    print("\nMake sure Excel is open with your SmartOptionChainExcel file!")
    input("Press Enter to start extraction...")
    
    extractor = EnhancedDateExtractor()
    dates = extractor.extract_all_dates()
    
    print("\n" + "="*70)
    print("EXTRACTION COMPLETE!")
    print("="*70)
    print("\nResults saved to: excel_dates_cache.json")
    print("\nYou can now use these dates in your main system.")
    print("="*70)