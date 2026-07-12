#!/usr/bin/env python3
"""
Backward-compatible entry point — delegates to run_pipeline.py.

Usage (unchanged):
    python meesho_monthly_update.py "MeeshoMasterfilesampleautomation.xlsx" \\
        --month "May'26" --raw-data raw_may26.csv
"""
from run_pipeline import main

if __name__ == "__main__":
    main()
