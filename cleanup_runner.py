#!/usr/bin/env python3
import os
import sys
import time

# Add the src directory to the Python path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

# Load environment variables FIRST
from dotenv import load_dotenv
load_dotenv()

# Import our logging service
from OLLibrary.utils.log_service import setup_logging, get_logger

# Set up logging
setup_logging(app_name="ACRA_CleanupRunner")
logger = get_logger(__name__)

if __name__ == "__main__":
    logger.info("Starting cleanup process")
    
    try:
        # Import and run the cleanup service
        from services import cleanup_orphaned_folders
        
        # Record the start time
        start_time = time.time()
        
        # Run the cleanup
        result = cleanup_orphaned_folders()
        
        # Calculate and log the execution time
        execution_time = time.time() - start_time
        logger.info(f"Cleanup completed successfully in {execution_time:.2f} seconds")
        logger.info(f"Cleanup results: {result}")
        
    except Exception as e:
        logger.error(f"Cleanup process failed: {str(e)}")
        sys.exit(1)
    
    logger.info("Cleanup process finished")
    sys.exit(0) 