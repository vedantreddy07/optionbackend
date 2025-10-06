# main.py - FIXED version with correct imports
import asyncio
import logging
from datetime import datetime, timedelta
from typing import Dict, Any, List
import time
from pathlib import Path
import json

from final_excel_handler import FinalExcelHandler
from kite_autologin import AutomatedDailyLogin
from date_extractor import EnhancedDateExtractor  # FIXED: Use correct class name

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class OptimizedOptionTradingSystem:
    """Trading system with dynamic date extraction from Excel"""
    
    def __init__(self):
        self.excel_handler = FinalExcelHandler()
        self.date_extractor = EnhancedDateExtractor()  # FIXED: Use correct class name
        self.token_data = None
        self.is_running = False
        self.initialization_complete = False
        self.API_KEY = "u2dbf0jrukko9urf"
        self.dropdown_options = None
        
    async def initialize_system_fast(self):
        """Initialize with auto-login and date extraction"""
        try:
            logger.info("Initializing system with auto-login...")
            start_time = time.time()
            
            # Step 1: Auto-login to get credentials
            login_system = AutomatedDailyLogin(api_key=self.API_KEY)
            
            if login_system.is_session_valid():
                logger.info("Using cached valid session")
                cache_file = Path("token_cache.json")
                with open(cache_file, 'r') as f:
                    self.token_data = json.load(f)
            else:
                logger.info("Performing fresh Kite login...")
                success = await login_system.perform_daily_login()
                
                if not success:
                    logger.error("Auto-login failed!")
                    # Use fallback credentials for testing
                    self.token_data = {
                        'user_id': 'JOL229',
                        'enc_token': 'fallback_token'
                    }
                else:
                    cache_file = Path("token_cache.json")
                    with open(cache_file, 'r') as f:
                        self.token_data = json.load(f)
            
            # Step 2: Extract dropdown options from Excel (including dates)
            logger.info("Extracting dropdown options from Excel...")
            self.dropdown_options = self.get_dropdown_options()
            
            self.is_running = True
            self.initialization_complete = True
            
            init_time = time.time() - start_time
            logger.info(f"System initialized in {init_time:.2f} seconds")
            
            # Log what was extracted
            logger.info(f"Available symbols: {self.dropdown_options.get('symbols', [])}")
            logger.info(f"Option expiry dates: {len(self.dropdown_options.get('option_expiry', []))} dates")
            logger.info(f"Future expiry dates: {len(self.dropdown_options.get('future_expiry', []))} dates")
            
            return True
            
        except Exception as e:
            logger.error(f"Initialization failed: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_dropdown_options(self) -> Dict[str, List[str]]:
        """Get dropdown options dynamically from Excel"""
        
        logger.info("Extracting dates from Excel file...")
        
        try:
            # Try to extract from Excel
            extracted = self.date_extractor.extract_all_dates()
            
            # Validate extraction
            if (extracted.get('symbols') and 
                extracted.get('option_expiry') and 
                extracted.get('future_expiry')):
                
                logger.info("Successfully extracted dates from Excel:")
                logger.info(f"  Symbols: {len(extracted['symbols'])}")
                logger.info(f"  Option Expiry: {len(extracted['option_expiry'])}")
                logger.info(f"  Future Expiry: {len(extracted['future_expiry'])}")
                
                # Save for reference
                self._save_extracted_dates(extracted)
                
                return extracted
            else:
                logger.warning("Extraction incomplete, using cached/fallback dates")
                return self._load_cached_or_fallback_dates()
        
        except Exception as e:
            logger.error(f"Date extraction error: {e}")
            import traceback
            traceback.print_exc()
            return self._load_cached_or_fallback_dates()
    
    def _save_extracted_dates(self, dates_dict: Dict):
        """Save extracted dates to cache file"""
        try:
            cache_file = Path("excel_dates_cache.json")
            with open(cache_file, 'w') as f:
                json.dump({
                    **dates_dict,
                    'extracted_at': datetime.now().isoformat()
                }, f, indent=2)
            logger.info("Dates cached to excel_dates_cache.json")
        except Exception as e:
            logger.warning(f"Failed to cache dates: {e}")
    
    def _load_cached_or_fallback_dates(self) -> Dict[str, List[str]]:
        """Load from cache or use fallback"""
        
        # Try cache first
        cache_file = Path("excel_dates_cache.json")
        if cache_file.exists():
            try:
                with open(cache_file, 'r') as f:
                    cached = json.load(f)
                
                # Check if cache is recent (less than 1 day old)
                extract_time = datetime.fromisoformat(cached['extracted_at'])
                if datetime.now() - extract_time < timedelta(days=1):
                    logger.info("Using cached dates from excel_dates_cache.json")
                    return {
                        'symbols': cached.get('symbols', []),
                        'option_expiry': cached.get('option_expiry', []),
                        'future_expiry': cached.get('future_expiry', [])
                    }
            except Exception as e:
                logger.debug(f"Cache load failed: {e}")
        
        # Fallback dates
        logger.warning("Using FALLBACK dates - may not match your Excel!")
        return {
            'symbols': ["NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY", "SENSEX"],
            'option_expiry': [
                "07-10-2025", "14-10-2025", "21-10-2025", "28-10-2025",
                "04-11-2025", "11-11-2025", "18-11-2025", "25-11-2025",
                "02-12-2025", "09-12-2025", "16-12-2025", "23-12-2025", "30-12-2025"
            ],
            'future_expiry': [
                "30-10-2025", "27-11-2025", "25-12-2025"
            ]
        }
    
    async def fetch_option_data_fast(self, symbol: str, option_expiry: str,
                                    future_expiry: str, chain_length: int) -> Dict[str, Any]:
        """
        Fetch LIVE data from Excel by:
        1. Entering credentials
        2. Setting inputs
        3. Clicking OPTION CHAIN button
        4. Extracting data
        """
        try:
            if not self.token_data:
                logger.error("No token data! Initialize system first.")
                return self._generate_fallback_data(symbol, option_expiry, 
                                                   future_expiry, chain_length)
            
            logger.info(f"Fetching live data: {symbol}, {option_expiry}, length={chain_length}")
            start_time = time.time()
            
            # This does the complete workflow including button click
            live_data = await self.excel_handler.fetch_live_data(
                token_data=self.token_data,
                symbol=symbol,
                option_expiry=option_expiry,
                future_expiry=future_expiry,
                chain_length=chain_length
            )
            
            fetch_time = time.time() - start_time
            logger.info(f"Data fetched in {fetch_time:.2f} seconds")
            logger.info(f"Data source: {live_data.get('data_source')}")
            
            # Validate data quality
            if live_data.get('data_source') == 'excel_live':
                logger.info("Live data successfully fetched!")
            else:
                logger.warning("Data may be stale - check Excel manually")
            
            return live_data
            
        except Exception as e:
            logger.error(f"Data fetch failed: {e}")
            import traceback
            traceback.print_exc()
            return self._generate_fallback_data(symbol, option_expiry, 
                                              future_expiry, chain_length)
    
    def _generate_fallback_data(self, symbol: str, option_expiry: str,
                               future_expiry: str, chain_length: int) -> Dict[str, Any]:
        """Generate fallback data if live fetch fails"""
        base_price = {'NIFTY': 19500, 'BANKNIFTY': 45000, 
                     'FINNIFTY': 19000, 'MIDCPNIFTY': 9500}.get(symbol, 19500)
        
        return {
            'symbol': symbol,
            'option_expiry': option_expiry,
            'future_expiry': future_expiry,
            'chain_length': chain_length,
            'spot_ltp': base_price,
            'market_data': {
                'spot_ltp': base_price,
                'spot_ltp_change': 0,
                'spot_ltp_change_pct': 0,
                'future_price': base_price,
                'pcr': 0,
                'max_pain': 0,
                'india_vix': 0
            },
            'option_chain': [],
            'data_source': 'fallback',
            'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
        }
    
    async def cleanup(self):
        """Cleanup"""
        self.is_running = False


# For backward compatibility
OptionTradingSystem = OptimizedOptionTradingSystem


# Test the system
if __name__ == "__main__":
    async def test():
        system = OptimizedOptionTradingSystem()
        await system.initialize_system_fast()
        
        print("\n" + "="*60)
        print("EXTRACTED DROPDOWN OPTIONS:")
        print("="*60)
        
        options = system.dropdown_options
        
        print("\nSymbols:")
        for sym in options.get('symbols', []):
            print(f"  - {sym}")
        
        print(f"\nOption Expiry Dates ({len(options.get('option_expiry', []))}):")
        for date in options.get('option_expiry', [])[:5]:  # Show first 5
            print(f"  - {date}")
        if len(options.get('option_expiry', [])) > 5:
            print(f"  ... and {len(options.get('option_expiry', [])) - 5} more")
        
        print(f"\nFuture Expiry Dates ({len(options.get('future_expiry', []))}):")
        for date in options.get('future_expiry', []):
            print(f"  - {date}")
        
        print("="*60)
        
        await system.cleanup()
    
    asyncio.run(test())