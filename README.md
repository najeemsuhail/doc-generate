# Document Generation from Excel - Setup Guide

## Overview
This project reads customer data from an Excel file and generates personalized Word documents (letters) for each customer.

## Features
- ✓ Reads customer data from Excel spreadsheet
- ✓ Generates personalized Word documents for each customer
- ✓ Customizes letter content based on customer status (Overdue, Delinquent, Pending, Active)
- ✓ Includes customer address information for mailing
- ✓ Professional letter formatting
- ✓ Batch processing of multiple customers

## Files Included
1. **create_sample_excel.py** - Creates a sample Excel file with test customer data
2. **generate_letters.py** - Main script that generates Word documents
3. **requirements.txt** - Python dependencies
4. **README.md** - This file

## Installation & Usage

### Step 1: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Create Sample Data (Optional)
```bash
python create_sample_excel.py
```
This creates `customers.xlsx` with 8 sample customers.

### Step 3: Run the Letter Generator
```bash
python generate_letters.py
```

The script will:
- Read `customers.xlsx`
- Generate Word documents for each customer
- Save them in the `output_letters/` folder
- Name each file as `Letter_[Customer_Name].docx`

## Excel File Format

Your Excel file should have these columns:
| Column | Type | Notes |
|--------|------|-------|
| Name | Text | Customer name |
| Address | Text | Street address |
| City | Text | City |
| State | Text | State abbreviation (e.g., IL) |
| Zip | Text | Zip code |
| Status | Text | Overdue, Delinquent, Pending, Active, etc. |
| Outstanding_Amount | Number | Amount owed (e.g., 1500.00) |

**Note:** Column names are case-sensitive and must match exactly.

## Customization

You can customize the following in `generate_letters.py`:

1. **Company Header** - Change the company details in the header section
2. **Letter Template** - Modify the body text based on status
3. **Letter Closing** - Update sender signature line
4. **Output Folder** - Change where letters are saved

## Example

Input (customers.xlsx):
| Name | Address | City | State | Zip | Status | Outstanding_Amount |
|------|---------|------|-------|-----|--------|-------------------|
| John Smith | 123 Main St | Springfield | IL | 62701 | Overdue | 1500.00 |

Output: Word document named `Letter_John_Smith.docx` with personalized content based on his status and outstanding amount.

## Status-Based Letter Content

- **Overdue/Delinquent** - Urgent payment request with multiple payment methods
- **Pending** - Reminder to process payment
- **Active** - General account status confirmation

## Requirements
- Python 3.7+
- openpyxl (Excel file handling)
- python-docx (Word document generation)
- pandas (Data manipulation)
