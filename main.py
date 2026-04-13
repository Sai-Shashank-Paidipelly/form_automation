"""
ES Windows Order Form Automation
Reads line items from an Excel file and fills them into the ES Windows order form.

Usage:
    1. Start Chrome with remote debugging (see setup.sh)
    2. Log into https://orders.eswindows.co manually
    3. Run: python3 main.py
"""

import sys
import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from form_filler import FormFiller, FormFillerError
import form_selectors as sel


def read_excel(file_path):
    """Read line items from the 'Data Entry' sheet, skipping empty rows."""
    wb = load_workbook(file_path, data_only=True)
    ws = wb["Data Entry"]

    headers = [cell.value for cell in ws[1]]
    rows = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        # Skip rows where Product Type is empty or "-SELECT-"
        product_type = row[0]
        if not product_type or str(product_type).strip() in ("", "-SELECT-"):
            continue
        row_dict = dict(zip(headers, row))
        rows.append(row_dict)

    return rows


def connect_to_chrome():
    """Attach to an existing Chrome session with remote debugging enabled."""
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    try:
        driver = webdriver.Chrome(options=options)
        print("Connected to Chrome successfully.")
        return driver
    except Exception as e:
        print(f"ERROR: Could not connect to Chrome: {e}")
        print()
        print("Make sure Chrome is running with remote debugging enabled:")
        print('  /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=9222')
        print()
        print("Or run: bash setup.sh")
        sys.exit(1)


def main():
    print("=" * 60)
    print("  ES Windows Order Form Automation")
    print("=" * 60)
    print()

    # Get order number
    order_number = input("Enter the order number: ").strip()
    if not order_number:
        print("Order number is required.")
        return

    # Get Excel file path
    excel_path = input("Enter the Excel file path (drag & drop the file here): ").strip()
    # Remove surrounding quotes if present (from drag & drop)
    excel_path = excel_path.strip("'\"")

    if not excel_path:
        print("Excel file path is required.")
        return

    # Read Excel data
    print(f"\nReading Excel file: {excel_path}")
    try:
        rows = read_excel(excel_path)
    except FileNotFoundError:
        print(f"ERROR: File not found: {excel_path}")
        return
    except Exception as e:
        print(f"ERROR: Could not read Excel file: {e}")
        return

    print(f"Found {len(rows)} line items to process.")

    if not rows:
        print("No data found in the 'Data Entry' sheet. Exiting.")
        return

    # Preview the data
    print("\nPreview of line items:")
    for i, row in enumerate(rows[:5]):
        print(f"  {i+1}. {row.get('Product Type', '?')} | {row.get('Model', '?')} | W:{row.get('Width', '?')} x H:{row.get('Height', '?')}")
    if len(rows) > 5:
        print(f"  ... and {len(rows) - 5} more")

    print()
    confirm = input(f"Proceed to fill {len(rows)} line items into order #{order_number}? (y/n): ").strip().lower()
    if confirm != "y":
        print("Cancelled.")
        return

    # Connect to Chrome
    print("\nConnecting to Chrome...")
    driver = connect_to_chrome()

    # Navigate to the order page
    url = sel.ORDER_URL.format(order_number=order_number)
    print(f"Navigating to: {url}")
    driver.get(url)
    time.sleep(4)

    # Initialize form filler
    filler = FormFiller(driver)

    # Process each line item
    success_count = 0
    error_count = 0

    for i, row in enumerate(rows):
        print(f"\n{'─' * 50}")
        print(f"Line Item {i+1}/{len(rows)}")
        print(f"  Product: {row.get('Product Type', '?')}")
        print(f"  Model:   {row.get('Model', '?')}")
        print(f"  Size:    {row.get('Width', '?')} x {row.get('Height', '?')}")

        try:
            filler.add_line_item(row)
            success_count += 1
            print(f"  >> CREATED successfully")
        except FormFillerError as e:
            print(f"\n  >> {e}")
            print(f"\n  Process terminated at line item {i+1}.")
            print(f"  Successfully created {success_count} out of {len(rows)} line items before termination.")
            return
        except Exception as e:
            error_count += 1
            print(f"  >> UNEXPECTED ERROR: {e}")
            print(f"\n  Process terminated at line item {i+1}.")
            print(f"  Successfully created {success_count} out of {len(rows)} line items before termination.")
            return

    # Summary
    print(f"\n{'=' * 50}")
    print(f"DONE!")
    print(f"  Successfully created: {success_count}")
    print(f"  Errors:              {error_count}")
    print(f"  Total attempted:     {success_count + error_count}")
    print(f"{'=' * 50}")


if __name__ == "__main__":
    main()
