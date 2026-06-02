import os
import re
from pathlib import Path
import pandas as pd

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None
    print(
        "Warning: 'pypdf' library not found. PDF files will be skipped. Run 'pip install pypdf' to enable."
    )


def consolidate_email_master(directory: str, master_filename: str):
    """
    Scans a directory for emails, merges them with a master file,
    removes duplicates, and saves the result.
    """
    # Standard email regex pattern
    email_regex = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

    # Use a set for O(1) lookups and automatic deduplication
    unique_emails = set()

    dir_path = Path(directory)
    master_path = dir_path / master_filename

    # 1. Load existing emails from the master file if it exists
    if master_path.exists():
        try:
            df_existing = pd.read_excel(master_path)
            if "Email" in df_existing.columns:
                # Extract emails, convert to lower, and drop NaNs
                existing = (
                    df_existing["Email"].dropna().astype(str).str.lower().tolist()
                )
                unique_emails.update(existing)
            print(
                f"Loaded {len(unique_emails)} unique emails from existing master Excel file."
            )
        except Exception as e:
            print(f"Error reading master file: {e}")

    # 2. Process all other files in the directory
    for file in dir_path.iterdir():
        # Skip the master file itself and directories
        if file.name == master_filename or not file.is_file():
            continue

        # Skip the script file itself if it's in the same directory
        if file.suffix == ".py":
            continue

        try:
            found = []
            ext = file.suffix.lower()

            if ext == ".pdf" and PdfReader:
                reader = PdfReader(file)
                text_content = ""
                for page in reader.pages:
                    text_content += page.extract_text() or ""
                found = email_regex.findall(text_content.lower())

            elif ext in [".xlsx", ".xls"]:
                df = pd.read_excel(file)
                # Convert all data in the excel to string and search for emails
                combined_text = (
                    df.astype(str).apply(lambda x: " ".join(x), axis=1).str.cat(sep=" ")
                )
                found = email_regex.findall(combined_text.lower())

            else:
                # Default handling for text, csv, or log files
                with open(file, "r", encoding="utf-8", errors="ignore") as f:
                    found = email_regex.findall(f.read().lower())

            if found:
                unique_emails.update(found)
                print(f"Processed '{file.name}': Found {len(found)} emails.")
        except Exception as e:
            print(f"Skipping '{file.name}' due to error: {e}")

    # 3. Save the consolidated list back to the master file
    sorted_emails = sorted(list(unique_emails))
    try:
        df_new = pd.DataFrame(sorted_emails, columns=["Email"])
        df_new.to_excel(master_path, index=False)
        print(f"\nTask Complete!")
        print(
            f"Master file '{master_filename}' now contains {len(sorted_emails)} unique emails."
        )
    except Exception as e:
        print(f"Failed to save master file: {e}")


if __name__ == "__main__":
    # Configure your directory and master file name here
    TARGET_DIRECTORY = "."
    MASTER_FILE = "master_emails.xlsx"

    consolidate_email_master(TARGET_DIRECTORY, MASTER_FILE)
