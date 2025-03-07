import os
import pandas as pd
from pathlib import Path

def normalize_expediente(text):
    """Normalize expediente string for comparison."""
    # Convert to string and normalize
    text = str(text).strip()
    
    # Remove prefix if present
    if text.startswith('Documentos-'):
        text = text[len('Documentos-'):]
        
    # Remove suffix if present
    if ' CON PASE' in text:
        text = text.replace(' CON PASE', '')
    if '.zip' in text:
        text = text.replace('.zip', '')
        
    # Handle special characters
    text = text.replace('#', '%')
    
    # Remove all whitespace
    text = "".join(text.split())
    
    return text

def main():
    # Get the absolute path to the downloads directory
    downloads_dir = Path("downloads").absolute()
    if not downloads_dir.exists():
        print(f"Error: Downloads directory not found at {downloads_dir}")
        return

    # Find Excel file in root directory
    excel_files = list(Path(".").glob("*.xlsx"))
    if not excel_files:
        print("Error: No Excel file found in root directory")
        return
    
    excel_path = excel_files[0]
    print(f"Using Excel file: {excel_path}")

    try:
        # Read Excel file, using column D for expedientes (no skipping rows)
        df = pd.read_excel(excel_path, usecols=[3])  # Column D is index 3 (0-based)
        expedientes = df.iloc[:, 0]  # Get the first (and only) column
        
        # Print detailed Excel information
        total_rows = len(expedientes)
        empty_rows = expedientes.isna().sum()
        valid_rows = total_rows - empty_rows
        print(f"\nExcel file details:")
        print(f"Total rows: {total_rows}")
        print(f"Empty rows: {empty_rows}")
        print(f"Valid expedientes: {valid_rows}")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Get list of downloaded files
    downloaded_files = list(downloads_dir.glob("*.zip"))
    
    # Track results
    missing_files = []
    extra_files = []
    found_files = []
    
    # Create dictionaries of normalized values for comparison
    normalized_downloads = {normalize_expediente(f.name): f.name for f in downloaded_files}
    
    # Include all rows, even empty ones, in the expedientes dictionary
    normalized_expedientes = {}
    for exp in expedientes:
        if pd.notna(exp):  # If not empty
            norm_exp = normalize_expediente(exp)
            normalized_expedientes[norm_exp] = exp
    
    print(f"\nChecking {len(normalized_expedientes)} expedientes against {len(downloaded_files)} downloaded files...")
    
    # Check for missing files
    for norm_exp, original_exp in normalized_expedientes.items():
        if norm_exp in normalized_downloads:
            found_files.append(normalized_downloads[norm_exp])
        else:
            missing_files.append(original_exp)
    
    # Check for extra files
    for norm_download, original_filename in normalized_downloads.items():
        if norm_download not in normalized_expedientes:
            extra_files.append(original_filename)
    
    # Print results
    print("\nResults:")
    print(f"Total rows in Excel (excluding header): {total_rows}")
    print(f"Valid expedientes in Excel: {valid_rows}")
    print(f"Total files in downloads: {len(downloaded_files)}")
    print(f"Files found: {len(found_files)}")
    
    if missing_files:
        print(f"\nMissing files ({len(missing_files)}):")
        for exp in missing_files:
            print(f"- {exp}")
    else:
        print("\nNo missing files!")
        
    if extra_files:
        print(f"\nExtra files in downloads ({len(extra_files)}):")
        for file in extra_files:
            print(f"- {file}")
    else:
        print("\nNo extra files!")

if __name__ == "__main__":
    main() 