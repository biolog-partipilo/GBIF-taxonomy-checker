# GBIF-taxonomy-checker
Python script to validate and update taxonomic names using GBIF, writing results directly into an Excel file

## Features

- Validates species names against GBIF
- Detects accepted names and synonyms
- Retrieves updated taxonomic names
- Writes results into Excel:
  - Column I → status (valid / syn / missing / error)
  - Column J → updated taxon name
  - Column K → authorship
- Progress bar in terminal
- Auto-save every 25 rows

## Requirements

- Python 3.10+
- requests
- openpyxl
- tqdm
- a file containing species list with a row for each name (default col. D)

Install dependencies:

```bash
pip install -r requirements.txt
```
## Usage

Edit the file path in the script:
```
  FILE_PATH = r"path\to\your\file.xlsx"
```
ensure to keep the file closed until the script has finished running.

Run the script: ```
  python taxa_check.py ```
  
Notes <br>
  GBIF matching is not perfect and may not resolve all synonym chains.<br>
  Some taxonomic updates (especially genus changes) may not be detected. <br>
  The script performs an additional lookup for synonyms to improve accuracy.<br>
  In some cases autorship isn't correctly separated from taxa name, 
  prentesis may be missing due to original formatting on GBIF 

Output <br>
The script updates the excel file directly:
  - Overwrites existing values
  - Saves progress every 25 rows (customisable)
  - separate species from author

  
## Limitations
  - Depends on GBIF backbone taxonomy
  - No caching implemented
  - Requires internet connection
  - tailored on .xlsx files (not tested with other formats)


## Citation
If you use this tool, please consider citing or acknowledging the repository.
