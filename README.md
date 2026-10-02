# Data Cleaning & Preparation (Project 1)

Cleaned an e-commerce orders dataset (1,200 rows, 14 columns) as part of the DecodeLabs Data Analytics internship.

## What I did
- Audited for duplicates, null values and format issues
- Converted dates to ISO 8601 (YYYY-MM-DD)
- Treated blank CouponCode as "NO COUPON" instead of imputing
- Verified: 0 duplicate IDs, 0 incorrectly formatted dates

## Files
- `Ayush1.xlsx` - raw data
- `Ayush1_cleaned.xlsx` - cleaned data, change log and verification sheets
- `change_log.pdf` - documented changes
- `clean_ayush1.py` - cleaning script

## How to run
pip install pandas openpyxl
python clean_ayush1.py
