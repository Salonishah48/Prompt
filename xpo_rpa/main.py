"""
XPO Billing RPA - Main Entry Point
Automates login to XPO portal, navigates to billing, sorts by latest date,
downloads bills in batches of 10, and generates an Excel report.
"""

import sys
import os
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from rpa.config import Config
from rpa.logger import setup_logger
from rpa.runner import RPARunner

def main():
    logger = setup_logger("XPO_RPA")
    logger.info("=" * 60)
    logger.info("XPO Billing RPA - Starting")
    logger.info("=" * 60)

    try:
        config = Config()
        config.validate()

        runner = RPARunner(config, logger)
        runner.run()

        logger.info("RPA completed successfully.")
    except Exception as e:
        logger.error(f"RPA failed: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
