from pathlib import Path
from openpyxl import Workbook

DATA_DIR = Path(__file__).parent / "data"
EXCEL_FILE = DATA_DIR / "expenses.xlsx"

HEADERS = ["ID", "Date", "Type", "Category", "Amount", "Description"]


def initialize_database():
    DATA_DIR.mkdir(exist_ok=True)

    if not EXCEL_FILE.exists():
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Transactions"
        sheet.append(HEADERS)
        workbook.save(EXCEL_FILE)
        workbook.close()

    print(f"Excel database ready: {EXCEL_FILE}")


if __name__ == "__main__":
    initialize_database()
