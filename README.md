# Monday.com Data Uploader

A Python-based GUI application for uploading and managing data in Monday.com boards. This application provides a user-friendly interface for uploading data from Excel/CSV files and manually editing data through a spreadsheet editor.

## Features

- Upload data from Excel (.xlsx, .xls) and CSV files
- Built-in spreadsheet editor for manual data entry
- Support for various Monday.com column types:
  - Text
  - Numbers
  - Email
  - Status
  - People (with dropdown selection)
  - Date
  - Phone
  - Long text
  - Board relations

## Recent Updates

### Email Handling Improvements
- Fixed email validation and formatting
- Proper JSON structure for email columns
- Improved error handling for invalid email formats

### Status Column Enhancements
- Added dropdown menu for status selection
- Dynamic loading of status options from board settings
- Support for both label and index-based status values
- Proper handling of board-specific status labels

### User Interface Improvements
- Spreadsheet editor window matches main window size
- Added dropdown menus for person and status columns
- Better error messages and validation feedback

## Requirements

- Python 3.x
- Required packages (install via pip):
  ```
  pip install -r requirements.txt
  ```

## Setup

1. Clone the repository:
   ```
   git clone https://github.com/fabiemdf/Monday-Uploader.git
   ```

2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   python main.py
   ```

## Usage

1. Enter your Monday.com API key
2. Click "Load Boards" to fetch available boards
3. Select a board from the dropdown
4. Choose one of the following options:
   - Select an Excel/CSV file to upload
   - Open the spreadsheet editor for manual data entry

### Using the Spreadsheet Editor

- Double-click on cells to edit
- Double-click on person columns to select users from a dropdown
- Double-click on status columns to select from available status options
- Use the buttons at the bottom to:
  - Load existing items
  - Save to Excel
  - Upload to Monday.com

## API Rate Limits

The application is subject to Monday.com's API rate limits. If you encounter a "Daily limit exceeded" error, you'll need to:
1. Wait until the next day for the limit to reset, or
2. Upgrade your Monday.com plan for higher limits

## Error Handling

The application includes error handling for:
- Invalid API keys
- Network issues
- Invalid data formats
- API rate limits
- Column type validation

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 