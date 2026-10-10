# Healthcare Pitch Deck Research Package

Company summaries for the 65 companies in `01_Database/Healthcare_Pitch_Deck_Database.xlsx` (the workbook supplied by the project owner).

| Folder | Content |
|---|---|
| 01_Database | The workbook the summaries are built from (unchanged) |
| 02_Company_Summaries | `Company_Summaries.docx` (Word, all companies), `TXT/` (one text file per company), `Company_Summaries_All.txt` |
| 03_Tools | `build_summaries.py`, `validate_database.py`, `scraper.py`, `common.py`, `company_descriptions.csv`, `deck_notes.csv`, `requirements.txt` |
| 04_Reports | `Workbook_Check_and_v2_Comparison.xlsx`: automated checks on the workbook and its differences from the earlier v2 file |

## How the summaries are built

- Every statement comes from the workbook. Figures are shown as recorded.
- Website links are read from the hyperlinks in the Website column.
- Where the workbook disagrees with itself, the summary says so and does not choose. In 61 of 65 rows the Amount Raised text differs from the Amount (USD M) column of the same row.
- `company_descriptions.csv` only rewords the one-line business descriptions into sentences. Companies without a row there get the workbook's own description.

## Deck content

`03_Tools/deck_notes.csv` (columns: Name, Topic, Text, Source) adds a 'Deck content' section to a company's summary. It currently holds sourced notes for Alan and Hinge Health, taken from the text of their deck pages (the slides there are images and could not be read). To add a deck, append rows with the topic (for example Problem, Solution, Market, Traction, Team, Ask), the text, and the source, then rebuild the summaries.

## Commands

Install: `pip install -r 03_Tools/requirements.txt` (OCR for `scraper.py` also needs `tesseract-ocr` and `poppler-utils`).

1. Rebuild the summaries after editing the workbook:
   `python 03_Tools/build_summaries.py 01_Database/Healthcare_Pitch_Deck_Database.xlsx --out 02_Company_Summaries`
2. Check the workbook: `python 03_Tools/validate_database.py 01_Database/Healthcare_Pitch_Deck_Database.xlsx`
   (add `--compare OTHER.xlsx --report report.xlsx` to list differences between two versions)
3. Collect deck pages and files (needs normal internet access; test with `--limit 5` first):
   `python 03_Tools/scraper.py --out ./healthcare_decks`

## Limits

- No original deck files are downloaded and no slide text has been extracted. The summaries are not deck transcripts.
- Funding details come from secondary directories and have not been checked against original decks.
- `scraper.py` has been tested only on sample HTML, not on the live sites.
