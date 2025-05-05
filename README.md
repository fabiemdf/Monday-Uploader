# Monday.com Uploader App

A desktop application for uploading and editing items on Monday.com boards using a spreadsheet-like interface. Built with Python, Tkinter, pandas, and tksheet.

## Features
- **Connect to Monday.com**: Enter your API key and select a board.
- **Spreadsheet Editor**: Edit board items in a familiar spreadsheet UI (powered by tksheet).
- **Bulk Upload**: Import data from Excel/CSV and upload as new items.
- **Edit Existing Items**: Load items from a board, edit them, and save changes back to Monday.com.
- **Column Type Handling**: Supports text, number, status, date, email, phone, board-relation, and more.
- **Save/Load Spreadsheet**: Save your work locally as Excel for backup or offline editing.

## Installation

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd monday_uploader
   ```

2. **Create a virtual environment (optional but recommended)**
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```
   If `requirements.txt` is missing, install manually:
   ```bash
   pip install pandas requests tksheet openpyxl
   ```

## Usage

1. **Run the app**
   ```bash
   python main.py
   ```

2. **Connect to Monday.com**
   - Enter your Monday.com API key.
   - Click "Load Boards" and select a board.

3. **Upload Data**
   - Click "Select Excel/CSV File" to import data and upload as new items.

4. **Edit Board Items**
   - Click "Open Spreadsheet Editor".
   - Use "Load Items" to fetch current board items into the spreadsheet.
   - Edit values directly in the grid.
   - Click "Upload to Monday.com" to save changes (existing items are updated, new rows are created).
   - Use "Save to Excel" to export your current sheet for backup.

## Development Notes

- **Tech Stack**: Python, Tkinter (GUI), pandas (data handling), requests (API), tksheet (spreadsheet UI).
- **API Integration**: Uses Monday.com GraphQL API for all board/item operations.
- **Column Mapping**: Dynamically fetches board columns and maps spreadsheet columns to Monday.com column IDs/types.
- **Editing**: When loading items, their IDs are tracked. On upload, existing items are updated via `change_multiple_column_values`, and new rows are created with `create_item`.
- **Special Columns**: Handles special types (status, date, email, phone, board-relation, etc.) with correct JSON formatting for Monday.com API.
- **Error Handling**: API and data errors are shown in popups and printed to the console for debugging.

## Troubleshooting

- **API Errors**: Ensure your API key is valid and has access to the selected board.
- **Column Mismatches**: The spreadsheet columns must match the board's columns by title for correct mapping.
- **Dependencies**: If you see import errors, ensure all dependencies are installed (see Installation).
- **GraphQL Errors**: If you see GraphQL validation errors, check for API changes or column type mismatches.
- **Large Boards**: For boards with many items/columns, loading and uploading may take longer.

## Contributing
Pull requests and suggestions are welcome! Please open an issue for bugs or feature requests.

## License
MIT License 