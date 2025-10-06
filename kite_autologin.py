# kite_auto_login.py - Enhanced automated KiteConnect login
import asyncio
import logging
from datetime import datetime, time as dt_time
from typing import Dict, Optional
import json
from pathlib import Path
import time

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    print("Error: selenium not available. Install: pip install selenium")

try:
    import win32com.client as win32
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("Warning: pywin32 not available")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class KiteAutoLogin:
    """Handles automated KiteConnect login and enctoken extraction"""
    
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.driver = None
        self.enctoken = None
        self.user_id = None
        
    def setup_driver(self) -> bool:
        """Setup Chrome driver with necessary options"""
        try:
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            # Enable logging to capture network requests
            options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
            
            self.driver = webdriver.Chrome(options=options)
            
            # Remove webdriver flag
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            return True
        except Exception as e:
            logger.error(f"Chrome setup failed: {e}")
            return False
    
    async def login_and_extract_token(self) -> Dict[str, str]:
        """
        Main method: Opens KiteConnect login, waits for user to login,
        then extracts enctoken from browser
        """
        if not SELENIUM_AVAILABLE:
            logger.error("Selenium not available")
            return {}
        
        try:
            # Setup browser
            if not self.setup_driver():
                return {}
            
            logger.info("Opening KiteConnect login page...")
            
            # Navigate to Kite login
            self.driver.get("https://kite.zerodha.com/")
            
            logger.info("Waiting for user to login (2 minutes timeout)...")
            logger.info("Please login with your credentials")
            
            # Wait for successful login (redirect to dashboard)
            try:
                WebDriverWait(self.driver, 120).until(
                    EC.url_contains("dashboard")
                )
                logger.info("Login successful! Extracting token...")
                
            except TimeoutException:
                logger.error("Login timeout - user did not login within 2 minutes")
                return {}
            
            # Give page time to load
            await asyncio.sleep(2)
            
            # Method 1: Extract from cookies (most reliable)
            token_data = self._extract_from_cookies()
            
            if token_data.get('enctoken'):
                logger.info("Token extracted from cookies")
                return token_data
            
            # Method 2: Extract from localStorage
            token_data = self._extract_from_localstorage()
            
            if token_data.get('enctoken'):
                logger.info("Token extracted from localStorage")
                return token_data
            
            # Method 3: Extract from network logs
            token_data = self._extract_from_network_logs()
            
            if token_data.get('enctoken'):
                logger.info("Token extracted from network logs")
                return token_data
            
            logger.error("Could not extract token from any source")
            return {}
            
        except Exception as e:
            logger.error(f"Login/extraction failed: {e}")
            return {}
        
        finally:
            # Keep browser open for 3 seconds so user can see success
            if self.driver:
                await asyncio.sleep(3)
                try:
                    self.driver.quit()
                except:
                    pass
    
    def _extract_from_cookies(self) -> Dict[str, str]:
        """Extract token from cookies - MOST RELIABLE"""
        try:
            cookies = self.driver.get_cookies()
            
            enctoken = None
            user_id = None
            
            for cookie in cookies:
                if cookie['name'] == 'enctoken':
                    enctoken = cookie['value']
                elif cookie['name'] == 'user_id':
                    user_id = cookie['value']
            
            if enctoken:
                logger.info(f"Found enctoken in cookies (length: {len(enctoken)})")
                return {
                    'enctoken': enctoken,
                    'user_id': user_id or 'UNKNOWN'
                }
        except Exception as e:
            logger.debug(f"Cookie extraction failed: {e}")
        
        return {}
    
    def _extract_from_localstorage(self) -> Dict[str, str]:
        """Extract token from localStorage"""
        try:
            enctoken = self.driver.execute_script("return localStorage.getItem('enctoken')")
            user_id = self.driver.execute_script("return localStorage.getItem('user_id')")
            
            if enctoken:
                return {
                    'enctoken': enctoken,
                    'user_id': user_id or 'UNKNOWN'
                }
        except Exception as e:
            logger.debug(f"localStorage extraction failed: {e}")
        
        return {}
    
    def _extract_from_network_logs(self) -> Dict[str, str]:
        """Extract token from network request logs"""
        try:
            logs = self.driver.get_log('performance')
            
            for log in logs:
                try:
                    log_data = json.loads(log['message'])
                    message = log_data.get('message', {})
                    
                    # Look for network response
                    if message.get('method') == 'Network.responseReceived':
                        response = message.get('params', {}).get('response', {})
                        headers = response.get('headers', {})
                        
                        # Check for Set-Cookie with enctoken
                        set_cookie = headers.get('set-cookie', '') or headers.get('Set-Cookie', '')
                        
                        if 'enctoken' in set_cookie:
                            # Parse enctoken from Set-Cookie
                            for part in set_cookie.split(';'):
                                if 'enctoken=' in part:
                                    enctoken = part.split('enctoken=')[1].split(';')[0]
                                    
                                    # Try to get user_id
                                    user_id = self.driver.execute_script("return localStorage.getItem('user_id')")
                                    
                                    return {
                                        'enctoken': enctoken,
                                        'user_id': user_id or 'UNKNOWN'
                                    }
                
                except Exception as e:
                    continue
        
        except Exception as e:
            logger.debug(f"Network log extraction failed: {e}")
        
        return {}

class ExcelTokenWriter:
    """Writes extracted token to Excel file"""
    
    def __init__(self, excel_path: str = "SmartOptionChainExcel_Zerodha.xlsm"):
        self.excel_path = Path(excel_path).resolve()
        self.excel_app = None
        self.workbook = None
        self.worksheet = None
    
    def write_token_to_excel(self, user_id: str, enctoken: str) -> bool:
        """Write user_id and enctoken to Excel"""
        if not WIN32_AVAILABLE:
            logger.warning("pywin32 not available, cannot write to Excel")
            logger.info(f"USER_ID: {user_id}")
            logger.info(f"ENCTOKEN: {enctoken}")
            return False
        
        if not self.excel_path.exists():
            logger.error(f"Excel file not found: {self.excel_path}")
            logger.info(f"USER_ID: {user_id}")
            logger.info(f"ENCTOKEN: {enctoken}")
            return False
        
        try:
            pythoncom.CoInitialize()
            
            logger.info("Opening Excel file...")
            
            # Try to connect to existing Excel
            try:
                self.excel_app = win32.GetActiveObject("Excel.Application")
                logger.info("Connected to existing Excel instance")
            except:
                self.excel_app = win32.Dispatch("Excel.Application")
                self.excel_app.Visible = True
                logger.info("Created new Excel instance")
            
            # Open workbook
            try:
                for wb in self.excel_app.Workbooks:
                    if wb.Name == self.excel_path.name:
                        self.workbook = wb
                        break
                
                if not self.workbook:
                    self.workbook = self.excel_app.Workbooks.Open(str(self.excel_path))
                
                self.worksheet = self.workbook.ActiveSheet
                
            except Exception as e:
                logger.error(f"Failed to open workbook: {e}")
                return False
            
            # Write credentials to cells
            # ADJUST THESE CELL REFERENCES TO MATCH YOUR EXCEL FILE
            user_id_cell = "F587"  # Change this to your actual cell
            enctoken_cell = "F615"  # Change this to your actual cell
            
            logger.info(f"Writing user_id to cell {user_id_cell}")
            self.worksheet.Range(user_id_cell).Value = user_id
            
            logger.info(f"Writing enctoken to cell {enctoken_cell}")
            self.worksheet.Range(enctoken_cell).Value = enctoken
            
            logger.info("Credentials written to Excel successfully!")
            
            # Save workbook
            self.workbook.Save()
            logger.info("Workbook saved")
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to write to Excel: {e}")
            return False
        
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

class AutomatedDailyLogin:
    """Main orchestrator for daily automated login"""
    
    def __init__(self, api_key: str, excel_path: str = "SmartOptionChainExcel_Zerodha.xlsm"):
        self.api_key = api_key
        self.kite_login = KiteAutoLogin(api_key)
        self.excel_writer = ExcelTokenWriter(excel_path)
    
    async def perform_daily_login(self) -> bool:
        """
        Complete daily login workflow:
        1. Open Kite login
        2. Wait for user to login
        3. Extract enctoken
        4. Write to Excel
        """
        logger.info("Starting automated daily login process...")
        
        # Step 1 & 2 & 3: Login and extract token
        token_data = await self.kite_login.login_and_extract_token()
        
        if not token_data or not token_data.get('enctoken'):
            logger.error("Failed to extract token")
            return False
        
        logger.info(f"Token extracted successfully")
        logger.info(f"   User ID: {token_data['user_id']}")
        logger.info(f"   Token: {token_data['enctoken'][:20]}...")
        
        # Step 4: Write to Excel
        success = self.excel_writer.write_token_to_excel(
            user_id=token_data['user_id'],
            enctoken=token_data['enctoken']
        )
        
        if not success:
            logger.error("Failed to write token to Excel")
            return False
        
        # Cache token for session
        self._cache_token(token_data)
        
        logger.info("Daily login completed successfully!")
        logger.info("Excel is ready to fetch data")
        
        return True
    
    def _cache_token(self, token_data: Dict):
        """Cache token for the session"""
        try:
            cache_file = Path("token_cache.json")
            with open(cache_file, 'w') as f:
                json.dump({
                    **token_data,
                    'timestamp': datetime.now().isoformat()
                }, f)
            logger.info("Token cached for session")
        except Exception as e:
            logger.warning(f"Failed to cache token: {e}")
    
    def is_session_valid(self) -> bool:
        """Check if current session is still valid (same day)"""
        try:
            cache_file = Path("token_cache.json")
            if not cache_file.exists():
                return False
            
            with open(cache_file, 'r') as f:
                cached = json.load(f)
            
            cache_time = datetime.fromisoformat(cached['timestamp'])
            now = datetime.now()
            
            # Check if it's the same day and before 5 PM
            if cache_time.date() == now.date() and now.time() < dt_time(17, 0):
                logger.info("Existing session is valid")
                return True
            
            return False
            
        except Exception as e:
            logger.debug(f"Session check failed: {e}")
            return False

# Main execution
async def main():
    """Main execution function"""
    
    # IMPORTANT: Replace with your actual API key
    API_KEY = "u2dbf0jrukko9urf"  # Replace this!
    
    login_system = AutomatedDailyLogin(
        api_key=API_KEY,
        excel_path="SmartOptionChainExcel_Zerodha.xlsm"
    )
    
    # Check if session is already valid
    if login_system.is_session_valid():
        logger.info("Session already valid for today")
        return
    
    # Perform daily login
    success = await login_system.perform_daily_login()
    
    if success:
        print("\n" + "="*60)
        print("DAILY LOGIN COMPLETED SUCCESSFULLY!")
        print("="*60)
        print("Your Excel file is now ready to fetch live data")
        print("Session valid until 5:00 PM today")
        print("="*60)
    else:
        print("\n" + "="*60)
        print("LOGIN FAILED")
        print("="*60)
        print("Please check the logs above for details")
        print("="*60)

if __name__ == "__main__":
    asyncio.run(main())