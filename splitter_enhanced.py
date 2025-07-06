#!/usr/bin/env python3

import sys
import os
import shutil
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pathlib import Path
import glob
import argparse


def find_column(worksheet, column_name):
    for col_idx, cell in enumerate(worksheet[1], start=1):
        if cell.value == column_name:
            return col_idx
    raise ValueError(f"Cannot find '{column_name}' column! Please check column name")


def copy_documents(source_dir, dest_dir, app_name):
    word_pattern = os.path.join(source_dir, f"{app_name}*.docx")
    word_files = glob.glob(word_pattern)
    
    pdf_pattern = os.path.join(source_dir, f"{app_name}*permission*.pdf")
    pdf_files = glob.glob(pdf_pattern)
    
    # For testing, also look for .txt files
    if not word_files:
        word_pattern = os.path.join(source_dir, f"{app_name}*.txt")
        word_files = [f for f in glob.glob(word_pattern) if "permission" not in f]
    
    if not pdf_files:
        pdf_pattern = os.path.join(source_dir, f"{app_name}*permission*.txt")
        pdf_files = glob.glob(pdf_pattern)
    
    copied_files = []
    
    for word_file in word_files:
        dest_path = os.path.join(dest_dir, os.path.basename(word_file))
        shutil.copy2(word_file, dest_path)
        copied_files.append(os.path.basename(word_file))
    
    for pdf_file in pdf_files:
        dest_path = os.path.join(dest_dir, os.path.basename(pdf_file))
        shutil.copy2(pdf_file, dest_path)
        copied_files.append(os.path.basename(pdf_file))
    
    return copied_files


def create_sharepoint_sharing_script(base_dir, reviewers):
    script_path = os.path.join(base_dir, "share_folders.ps1")
    
    with open(script_path, 'w', encoding='utf-8') as f:
        f.write("# PowerShell script to share folders on SharePoint\n")
        f.write("# Run this script after uploading folders to SharePoint\n\n")
        f.write("$siteUrl = Read-Host 'Enter SharePoint site URL'\n")
        f.write("$baseFolder = Read-Host 'Enter base folder path on SharePoint'\n\n")
        f.write("Connect-PnPOnline -Url $siteUrl -UseWebLogin\n\n")
        
        for reviewer in reviewers:
            reviewer_name = str(reviewer).strip()
            f.write(f"# Share folder for {reviewer_name}\n")
            f.write(f"$folderPath = Join-Path $baseFolder '{reviewer_name}'\n")
            f.write(f"$userEmail = Read-Host 'Enter email for {reviewer_name}'\n")
            f.write(f"Set-PnPFolderPermission -List 'Documents' -Identity $folderPath -User $userEmail -AddRole 'Edit'\n")
            f.write(f"Write-Host 'Shared folder for {reviewer_name} with Edit permissions'\n\n")
    
    return script_path


def split_excel_enhanced(file_path, app_name):
    if not os.path.exists(file_path):
        print(f"Error: File not found {file_path}")
        sys.exit(1)
    
    print(f"Reading file: {file_path}")
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        print(f"Failed to read Excel: {e}")
        sys.exit(1)
    
    if 'Reviewer' not in df.columns:
        print("Error: Cannot find 'Reviewer' column in Excel")
        sys.exit(1)
    
    reviewers = df['Reviewer'].dropna().unique().tolist()
    print(f"Found {len(reviewers)} reviewers: {', '.join(str(r) for r in reviewers)}")
    
    base_dir = os.path.dirname(file_path)
    app_folder = os.path.join(base_dir, app_name)
    os.makedirs(app_folder, exist_ok=True)
    print(f"Created application folder: {app_folder}")
    
    base_name = os.path.basename(file_path)
    
    for reviewer in reviewers:
        reviewer_name = str(reviewer).strip()
        
        reviewer_folder = os.path.join(app_folder, reviewer_name)
        os.makedirs(reviewer_folder, exist_ok=True)
        
        dst_path = os.path.join(reviewer_folder, base_name)
        
        wb = load_workbook(file_path)
        ws = wb.active
        
        try:
            reviewer_col = find_column(ws, 'Reviewer')
            
            max_row = ws.max_row
            max_col = ws.max_column
            
            filter_range = f"A1:{get_column_letter(max_col)}{max_row}"
            ws.auto_filter.ref = filter_range
            
            ws.auto_filter.add_filter_column(reviewer_col - 1, [reviewer_name])
            
            wb.save(dst_path)
            print(f"✓ Created filtered Excel for {reviewer_name}")
            
            copied_docs = copy_documents(base_dir, reviewer_folder, app_name)
            if copied_docs:
                print(f"  ✓ Copied documents: {', '.join(copied_docs)}")
            
        except Exception as e:
            print(f"✗ Error processing {reviewer_name}: {e}")
        finally:
            wb.close()
    
    script_path = create_sharepoint_sharing_script(app_folder, reviewers)
    print(f"\n✓ Created SharePoint sharing script: {script_path}")
    
    print("\nProcessing complete!")
    print(f"All files have been created in: {app_folder}")
    print("\nNext steps:")
    print("1. Upload the entire folder structure to SharePoint")
    print("2. Run the PowerShell script 'share_folders.ps1' to set permissions")
    print("3. The script will prompt for reviewer email addresses")


def main():
    parser = argparse.ArgumentParser(description='Split Excel by reviewer with enhanced features')
    parser.add_argument('excel_file', help='Path to the Excel file')
    parser.add_argument('app_name', help='Application name for the main folder')
    
    args = parser.parse_args()
    
    split_excel_enhanced(args.excel_file, args.app_name)


if __name__ == "__main__":
    main()