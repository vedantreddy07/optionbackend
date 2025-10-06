"""
FastAPI Server for Option Trading System
Essential APIs: Initialize, Dropdown Options, Fetch Data
"""

import logging
import asyncio
from datetime import datetime
from typing import Dict, Any
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
import uvicorn

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("option_trading_api")

# Initialize FastAPI
app = FastAPI(
    title="Option Trading System API",
    description="Real-time option chain data from Excel",
    version="1.0.0"
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global system instance
trading_system = None


# Request Models
class FetchDataRequest(BaseModel):
    symbol: str = Field(..., description="Trading symbol (e.g., NIFTY, BANKNIFTY)")
    option_expiry: str = Field(..., description="Option expiry date (DD-MM-YYYY)")
    future_expiry: str = Field(..., description="Future expiry date (DD-MM-YYYY)")
    chain_length: int = Field(20, ge=5, le=50, description="Number of strikes to fetch")


# Startup
@app.on_event("startup")
async def startup_event():
    """Initialize system on startup"""
    global trading_system
    logger.info("=" * 70)
    logger.info("OPTION TRADING SYSTEM API SERVER STARTING")
    logger.info("=" * 70)
    
    try:
        from main import OptimizedOptionTradingSystem
        trading_system = OptimizedOptionTradingSystem()
        logger.info("Trading system instance created")
        
        # Auto-initialize in background
        asyncio.create_task(initialize_system_background())
        
    except Exception as e:
        logger.error(f"Startup failed: {e}")
        trading_system = None


async def initialize_system_background():
    """Initialize system in background"""
    global trading_system
    try:
        if trading_system:
            success = await trading_system.initialize_system_fast()
            if success:
                logger.info("System initialized successfully")
            else:
                logger.warning("System initialization incomplete")
    except Exception as e:
        logger.error(f"Background initialization failed: {e}")


@app.on_event("shutdown")
async def shutdown_event():
    """Cleanup on shutdown"""
    global trading_system
    logger.info("Shutting down...")
    if trading_system:
        try:
            await trading_system.cleanup()
        except:
            pass


# API 1: Initialize System
@app.post("/api/initialize")
async def initialize_system():
    """
    Initialize the trading system
    - Performs Kite login
    - Extracts dates from Excel
    - Prepares system for data fetching
    
    Returns:
        JSON with status and dropdown options
    """
    global trading_system
    
    if not trading_system:
        raise HTTPException(status_code=500, detail="System not available")
    
    if trading_system.initialization_complete:
        return {
            "status": "success",
            "message": "System already initialized",
            "dropdown_options": trading_system.dropdown_options,
            "timestamp": datetime.now().isoformat()
        }
    
    try:
        logger.info("API: Initialize request received")
        success = await trading_system.initialize_system_fast()
        
        if success:
            return {
                "status": "success",
                "message": "System initialized successfully",
                "dropdown_options": trading_system.dropdown_options,
                "timestamp": datetime.now().isoformat()
            }
        else:
            raise HTTPException(
                status_code=500,
                detail="System initialization failed"
            )
            
    except Exception as e:
        logger.error(f"Initialization error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# API 2: Get Dropdown Options
@app.get("/api/dropdown-options")
async def get_dropdown_options():
    """
    Get available dropdown options
    
    Returns:
        JSON with symbols, option_expiry dates, future_expiry dates
    """
    global trading_system
    
    if not trading_system:
        raise HTTPException(status_code=500, detail="System not available")
    
    if not trading_system.initialization_complete:
        raise HTTPException(
            status_code=400,
            detail="System not initialized. Call /api/initialize first"
        )
    
    try:
        options = trading_system.dropdown_options or trading_system.get_dropdown_options()
        
        return {
            "status": "success",
            "data": {
                "symbols": options.get('symbols', []),
                "option_expiry": options.get('option_expiry', []),
                "future_expiry": options.get('future_expiry', [])
            },
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error fetching dropdown options: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# API 3: Fetch Option Data
@app.post("/api/fetch-option-data")
async def fetch_option_data(request: FetchDataRequest):
    """
    Fetch live option chain data from Excel
    
    Process:
    1. Validates inputs
    2. Writes credentials to Excel
    3. Sets parameters in Excel
    4. Clicks "OPTION CHAIN" button
    5. Waits for data refresh
    6. Extracts ALL columns and returns as JSON
    
    Returns:
        Complete JSON with all 20 columns (8 call + 1 strike + 8 put + 3 interpretations)
    """
    global trading_system
    
    if not trading_system:
        raise HTTPException(status_code=500, detail="System not available")
    
    if not trading_system.initialization_complete:
        raise HTTPException(
            status_code=400,
            detail="System not initialized. Call /api/initialize first"
        )
    
    try:
        logger.info(f"API: Fetch request - {request.symbol} {request.option_expiry}")
        
        import time
        start_time = time.time()
        
        # Fetch data from Excel
        data = await trading_system.fetch_option_data_fast(
            symbol=request.symbol,
            option_expiry=request.option_expiry,
            future_expiry=request.future_expiry,
            chain_length=request.chain_length
        )
        
        fetch_time = time.time() - start_time
        
        # Return complete JSON response
        response = {
            "status": "success",
            "message": "Option data fetched successfully",
            "timestamp": datetime.now().isoformat(),
            "fetch_time_seconds": round(fetch_time, 2),
            "request_params": {
                "symbol": request.symbol,
                "option_expiry": request.option_expiry,
                "future_expiry": request.future_expiry,
                "chain_length": request.chain_length
            },
            "data": {
                # Basic info
                "symbol": data.get('symbol', ''),
                "option_expiry": data.get('option_expiry', ''),
                "future_expiry": data.get('future_expiry', ''),
                "data_source": data.get('data_source', 'unknown'),
                "timestamp": data.get('timestamp', ''),
                
                # Market summary metrics
                "market_data": {
                    "spot_ltp": float(data.get('spot', {}).get('spot_ltp', 0)),
                    "spot_ltp_change": float(data.get('spot', {}).get('spot_ltp_change', 0)),
                    "spot_ltp_change_pct": float(data.get('spot', {}).get('spot_ltp_change_pct', 0)),
                    "future_price": float(data.get('future', {}).get('future_price', 0)),
                    "pcr": float(data.get('pcr', 0)),
                    "max_pain": float(data.get('max_pain', 0)),
                    "india_vix": float(data.get('india_vix', 0))
                },
                
                # Trading signals
                "signals": {
                    "intraday": data.get('signals', {}).get('intraday', ''),
                    "weekly": data.get('signals', {}).get('weekly', '')
                },
                
                # Complete option chain with ALL columns
                "option_chain": [
                    {
                        "strike": float(strike.get('strike', 0)),
                        "call": {
                            "interpretation": str(strike.get('call', {}).get('interpretation', '')),
                            "avg_price": float(strike.get('call', {}).get('avg_price', 0)),
                            "iv": float(strike.get('call', {}).get('iv', 0)),
                            "oi_change": float(strike.get('call', {}).get('oi_change', 0)),
                            "oi": float(strike.get('call', {}).get('oi', 0)),
                            "volume": float(strike.get('call', {}).get('volume', 0)),
                            "ltp_change": float(strike.get('call', {}).get('ltp_change', 0)),
                            "ltp": float(strike.get('call', {}).get('ltp', 0))
                        },
                        "put": {
                            "ltp": float(strike.get('put', {}).get('ltp', 0)),
                            "ltp_change": float(strike.get('put', {}).get('ltp_change', 0)),
                            "volume": float(strike.get('put', {}).get('volume', 0)),
                            "oi": float(strike.get('put', {}).get('oi', 0)),
                            "oi_change": float(strike.get('put', {}).get('oi_change', 0)),
                            "iv": float(strike.get('put', {}).get('iv', 0)),
                            "avg_price": float(strike.get('put', {}).get('avg_price', 0)),
                            "interpretation": str(strike.get('put', {}).get('interpretation', ''))
                        }
                    }
                    for strike in data.get('option_chain', [])
                ]
            }
        }
        
        logger.info(f"Data fetched in {fetch_time:.2f}s - {len(data.get('option_chain', []))} strikes")
        
        return JSONResponse(content=response)
        
    except Exception as e:
        logger.error(f"Fetch error: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))


# Health check endpoint
@app.get("/health")
async def health_check():
    """Health check endpoint"""
    system_ready = trading_system is not None and trading_system.initialization_complete
    
    return {
        "status": "healthy" if system_ready else "initializing",
        "system_ready": system_ready,
        "timestamp": datetime.now().isoformat()
    }


# Run server
if __name__ == "__main__":
    logger.info("Starting FastAPI server on http://0.0.0.0:8000")
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        log_level="info"
    )