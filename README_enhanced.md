# Enhanced Excel Splitter for SharePoint

An advanced tool that splits Excel files by reviewer, creates organized folder structures, copies related documents, and prepares SharePoint sharing scripts.

## Features

- Creates application-specific folder structure
- Splits Excel file by "Reviewer" column with filtered views
- Automatically extracts email addresses from "Email Address" column (if available)
- Copies Word documents and permission PDFs to each reviewer folder
- Generates PowerShell script for SharePoint permissions with pre-filled emails
- Maintains Excel formatting with AutoFilter

## Installation

1. Ensure Python 3.9+ is installed
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python splitter_enhanced.py <excel_file> <application_name>
```

### Example

```bash
python splitter_enhanced.py "user_listing.xlsx" "MyApp"
```

## Output Structure

```
Base Directory/
├── MyApp/                              # Application folder
│   ├── John Doe/                       # Reviewer folder
│   │   ├── user_listing.xlsx           # Filtered Excel (John's records only)
│   │   ├── MyApp_Guide.docx           # Copied Word document
│   │   └── MyApp_permission_form.pdf   # Copied permission PDF
│   ├── Jane Smith/
│   │   ├── user_listing.xlsx
│   │   ├── MyApp_Guide.docx
│   │   └── MyApp_permission_form.pdf
│   └── share_folders.ps1               # SharePoint sharing script
```

## Document Copying Rules

- **Word Documents**: Copies all `.docx` files starting with the application name
- **PDFs**: Copies all `.pdf` files with both application name and "permission" in filename

## SharePoint Integration

### 1. Upload to SharePoint
Upload the entire application folder structure to your SharePoint document library.

### 2. Set Permissions
Run the generated PowerShell script:

```powershell
.\share_folders.ps1
```

The script will:
- Connect to your SharePoint site
- Use email addresses from Excel file (if "Email Address" column exists)
- Prompt for missing email addresses only
- Grant "Edit" permissions to their respective folders

### 3. Manual Alternative
If you prefer manual sharing:
1. Right-click each reviewer folder in SharePoint
2. Select "Share"
3. Enter reviewer's email
4. Set permission to "Can edit"
5. Send invitation

## Requirements

- Excel file must contain a "Reviewer" column
- (Optional) "Email Address" column for automatic SharePoint sharing
- Word documents should be named: `<AppName>*.docx`
- Permission PDFs should be named: `<AppName>*permission*.pdf`
- PowerShell with PnP module for SharePoint sharing (optional)

## PowerShell Setup (for automated sharing)

Install SharePoint PnP PowerShell module:
```powershell
Install-Module -Name PnP.PowerShell
```

## Troubleshooting

### Missing Reviewer Column
Ensure your Excel file has a column named exactly "Reviewer" (case-sensitive).

### Documents Not Copied
Check that Word/PDF files follow the naming convention and exist in the same directory as the Excel file.

### SharePoint Permissions Error
- Ensure you have site owner/admin permissions
- Verify reviewer email addresses are correct
- Check that folders are uploaded to SharePoint before running the script

## Notes

- Filtered Excel files maintain links to original data
- Each reviewer only sees their assigned records
- Original Excel file is not modified
- All timestamps and metadata are preserved