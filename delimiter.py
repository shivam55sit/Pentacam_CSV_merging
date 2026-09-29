"""
Detect the delimiter used in each CSV file in a folder.
"""
 
from pathlib import Path
 
# ============ CONFIG ============
INPUT_FOLDER = r"C:\HYD Extracted Pentacam\ADITHYA TALLENT"
 
# Candidate delimiters to try (auto-detected per file)
CSV_DELIMITER_CANDIDATES = [";", ",", "\t"]
# =================================
 
 
def sniff_delimiter(first_line: str) -> str:
    """Pick whichever candidate delimiter appears most often in the header line."""
    counts = {d: first_line.count(d) for d in CSV_DELIMITER_CANDIDATES}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else CSV_DELIMITER_CANDIDATES[0]
 
 
def main():
    input_dir = Path(INPUT_FOLDER)
    if not input_dir.exists():
        print(f"ERROR: Folder not found: {INPUT_FOLDER}")
        return
 
    csv_files = [
        f for f in sorted(input_dir.iterdir())
        if f.suffix.lower() == ".csv" and not f.name.startswith("~$")
    ]
 
    if not csv_files:
        print(f"No CSV files found in {INPUT_FOLDER}")
        return
 
    for f in csv_files:
        with open(f, "r", encoding="utf-8-sig", errors="replace") as fh:
            first_line = fh.readline()
        delim = sniff_delimiter(first_line)
        display = {";": "semicolon (;)", ",": "comma (,)", "\t": "tab"}.get(delim, delim)
        print(f"{f.name}: {display}")
 
 
if __name__ == "__main__":
    main()