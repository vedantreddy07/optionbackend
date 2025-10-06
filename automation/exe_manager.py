import subprocess
import psutil
import time
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

class ExeManager:
    def __init__(self, exe_path: str = "SmartOptionChainExcel.exe"):
        self.exe_path = Path(exe_path).resolve()
        self.process = None
        
    def start_exe(self) -> bool:
        """Start the EXE file and keep it running"""
        try:
            if not self.exe_path.exists():
                logger.error(f"EXE file not found: {self.exe_path}")
                return False
            
            # Start the process
            self.process = subprocess.Popen(
                str(self.exe_path),
                cwd=self.exe_path.parent,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            
            # Wait a bit for the process to start
            time.sleep(3)
            
            # Check if process is still running
            if self.process.poll() is None:
                logger.info(f"EXE started successfully with PID: {self.process.pid}")
                return True
            else:
                logger.error("EXE failed to start or crashed immediately")
                return False
                
        except Exception as e:
            logger.error(f"Error starting EXE: {e}")
            return False
    
    def is_running(self) -> bool:
        """Check if EXE is still running"""
        if self.process is None:
            return False
        return self.process.poll() is None
    
    def stop_exe(self):
        """Stop the EXE process"""
        if self.process and self.is_running():
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()
            logger.info("EXE process stopped")