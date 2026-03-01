[README.md](https://github.com/user-attachments/files/25660475/README.md)
# Grant Liquidation Form

A Google Apps Script-based web application for managing grant liquidation requests with PDF generation.

## Features

- **Form Submission**: Enter branch name, date, budget purpose, and withdrawal amount
- **Item Table**: Add multiple line items with date, description, and amount
- **Auto-sorting**: Table rows automatically arrange by date (oldest to newest)
- **Date Validation**: Prevents future dates with error message
- **Automatic Calculations**: 
  - Total amount calculation
  - Excess/deficit calculation (Withdrawal - Total)
- **PDF Generation**: Creates a formatted PDF document with all submitted data
- **Data Storage**: Saves all submissions to a Google Spreadsheet organized by branch

## Setup

1. Create a Google Spreadsheet to store data
2. Create a Google Doc template with placeholders:
   - `{{BRANCH}}`
   - `{{DATE}}`
   - `{{BUDGET}}`
   - `{{WITHDRAWAL}}`
   - `{{TOTAL AMOUNT}}`
   - `{{EXCESS}}`
   - `{{ITEMS}}`
3. Update `Code.js` with your template and spreadsheet IDs
4. Deploy as Web App

## Usage

1. Fill in the form fields (Name/Branch, Date Created, Purpose for Budget, Amount of Withdrawal)
2. Add items to the table with date, description, and amount
3. Dates are automatically sorted from oldest to newest
4. View total amount and excess/deficit
5. Submit to generate PDF and save to spreadsheet
